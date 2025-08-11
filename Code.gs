/**
 * Brevo Auto Import Gmail
 * Automatically imports email addresses from Gmail to Brevo CRM
 * 
 * This script:
 * 1. Reads emails with "Brevo Import" label from Gmail
 * 2. Extracts email addresses from email bodies
 * 3. Imports contacts to Brevo CRM
 * 4. Manages Gmail labels based on import results
 */

// TEMPORARY API KEY - DELETE THIS LINE AFTER RUNNING setupApiKey()
const TEMP_API_KEY = 'xkeysib-xxxxxx';

// Constants
const BREVO_API_ENDPOINT = 'https://api.brevo.com/v3/contacts';
const BREVO_API_KEY_PROPERTY = 'BREVO_API_KEY';
const IMPORT_LABEL = 'Brevo/Import';
const SUCCESS_LABEL = 'Brevo/Success';
const ERROR_LABEL = 'Brevo/Error';
const BATCH_SIZE = 10; // Process emails in batches to avoid timeouts
const MAX_PROCESSING_TIME = 4 * 60 * 1000; // 4 minutes max execution time
const BREVO_LIST_ID = 14; // List ID to add contacts to

/**
 * Main function to process emails and import contacts to Brevo
 */
function processBrevoImports() {
  const startTime = Date.now();
  
  try {
    console.log('Starting Brevo import process...');
    
    // Validate API key setup
    if (!getBrevoApiKey()) {
      throw new Error('Brevo API key not configured. Please run setupApiKey() first.');
    }
    
    // Get or create required labels
    const labels = setupLabels();
    
    // Search for emails with import label
    const query = `label:"${IMPORT_LABEL}"`;
    const threads = GmailApp.search(query, 0, 50); // Limit to 50 threads
    
    if (threads.length === 0) {
      console.log('No emails found with Brevo Import label');
      return;
    }
    
    console.log(`Found ${threads.length} thread(s) to process`);
    
    let processedCount = 0;
    let successCount = 0;
    let errorCount = 0;
    
    // Process threads in batches
    for (let i = 0; i < threads.length; i += BATCH_SIZE) {
      // Check execution time
      if (Date.now() - startTime > MAX_PROCESSING_TIME) {
        console.log('Approaching execution time limit, stopping processing');
        break;
      }
      
      const batch = threads.slice(i, i + BATCH_SIZE);
      const batchResults = processBatch(batch, labels);
      
      processedCount += batchResults.processed;
      successCount += batchResults.success;
      errorCount += batchResults.errors;
      
      // Add small delay between batches to respect rate limits
      if (i + BATCH_SIZE < threads.length) {
        Utilities.sleep(1000); // 1 second delay
      }
    }
    
    console.log(`Processing complete. Processed: ${processedCount}, Success: ${successCount}, Errors: ${errorCount}`);
    
  } catch (error) {
    console.error('Error in processBrevoImports:', error);
    throw error;
  }
}

/**
 * Process a batch of email threads
 */
function processBatch(threads, labels) {
  let processed = 0;
  let success = 0;
  let errors = 0;
  
  threads.forEach(thread => {
    try {
      const messages = thread.getMessages();
      const emailAddresses = new Set(); // Use Set to avoid duplicates
      
      // Extract email addresses from all messages in the thread
      messages.forEach(message => {
        const body = message.getPlainBody();
        const htmlBody = message.getBody();
        
        // Extract emails from both plain text and HTML body
        const plainEmails = extractEmailAddresses(body);
        const htmlEmails = extractEmailAddresses(htmlBody);
        
        plainEmails.forEach(email => emailAddresses.add(email));
        htmlEmails.forEach(email => emailAddresses.add(email));
      });
      
      if (emailAddresses.size === 0) {
        console.log(`No email addresses found in thread: ${thread.getFirstMessageSubject()}`);
        labels.errorLabel.addToThread(thread);
        labels.importLabel.removeFromThread(thread);
        errors++;
      } else {
        // Import contacts to Brevo
        const importResults = importContactsToBrevo(Array.from(emailAddresses));
        
        if (importResults.success) {
          console.log(`Successfully imported ${emailAddresses.size} contact(s) from thread: ${thread.getFirstMessageSubject()}`);
          labels.successLabel.addToThread(thread);
          labels.importLabel.removeFromThread(thread);
          success++;
        } else {
          console.error(`Failed to import contacts from thread: ${thread.getFirstMessageSubject()}. Error: ${importResults.error}`);
          labels.errorLabel.addToThread(thread);
          labels.importLabel.removeFromThread(thread);
          errors++;
        }
      }
      
      processed++;
      
    } catch (error) {
      console.error(`Error processing thread: ${thread.getFirstMessageSubject()}`, error);
      labels.errorLabel.addToThread(thread);
      labels.importLabel.removeFromThread(thread);
      errors++;
      processed++;
    }
  });
  
  return { processed, success, errors };
}

/**
 * Extract email addresses from text using regex patterns
 */
function extractEmailAddresses(text) {
  if (!text) return [];
  
  const emailAddresses = [];
  
  // Enhanced regex patterns for different email formats
  const patterns = [
    // Standard email format: name@domain.com
    /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g,
    // Email with display name: "John Doe" <john@example.com>
    /"?([^"<>\r\n]+)"?\s*<([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})>/g,
    // Email with display name: John Doe <john@example.com>
    /([^<>\r\n]+)\s*<([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})>/g
  ];
  
  patterns.forEach(pattern => {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      if (match.length === 1) {
        // Simple email pattern
        const email = match[0].toLowerCase();
        const name = extractNameFromEmail(email);
        emailAddresses.push({ email, name });
      } else if (match.length === 3) {
        // Email with display name pattern
        const name = match[1].trim().replace(/"/g, '');
        const email = match[2].toLowerCase();
        emailAddresses.push({ email, name });
      }
    }
  });
  
  // Remove duplicates based on email address
  const uniqueEmails = [];
  const seenEmails = new Set();
  
  emailAddresses.forEach(contact => {
    if (!seenEmails.has(contact.email)) {
      seenEmails.add(contact.email);
      uniqueEmails.push(contact);
    }
  });
  
  return uniqueEmails;
}

/**
 * Extract a name from an email address (best effort)
 */
function extractNameFromEmail(email) {
  const localPart = email.split('@')[0];
  
  // Handle common patterns
  if (localPart.includes('.')) {
    return localPart.split('.').map(part => 
      part.charAt(0).toUpperCase() + part.slice(1)
    ).join(' ');
  } else if (localPart.includes('_')) {
    return localPart.split('_').map(part => 
      part.charAt(0).toUpperCase() + part.slice(1)
    ).join(' ');
  } else {
    return localPart.charAt(0).toUpperCase() + localPart.slice(1);
  }
}

/**
 * Import contacts to Brevo CRM
 */
function importContactsToBrevo(contacts) {
  const apiKey = getBrevoApiKey();
  
  if (!apiKey) {
    return { success: false, error: 'API key not configured' };
  }
  
  try {
    const results = [];
    
    // Import contacts one by one to handle individual errors
    contacts.forEach(contact => {
      try {
        const payload = {
          email: contact.email,
          attributes: {
            FIRSTNAME: contact.name || '',
            LASTNAME: ''
          },
          listIds: [BREVO_LIST_ID], // Add contact to list #14
          updateEnabled: true // Update if contact already exists
        };
        
        // Split name into first and last name if available
        if (contact.name && contact.name.includes(' ')) {
          const nameParts = contact.name.split(' ');
          payload.attributes.FIRSTNAME = nameParts[0];
          payload.attributes.LASTNAME = nameParts.slice(1).join(' ');
        }
        
        const response = UrlFetchApp.fetch(BREVO_API_ENDPOINT, {
          method: 'POST',
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'api-key': apiKey
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        
        const responseCode = response.getResponseCode();
        
        if (responseCode === 201 || responseCode === 204) {
          console.log(`Successfully imported contact: ${contact.email}`);
          results.push({ email: contact.email, success: true });
        } else if (responseCode === 400) {
          // Contact might already exist - try to update them to add to list
          let responseData = '';
          try {
            const responseText = response.getContentText();
            if (responseText) {
              responseData = JSON.parse(responseText);
            }
          } catch (parseError) {
            responseData = `Response code: ${responseCode}`;
          }
          
          // If contact exists, update them to add to the list
          if (responseData.code === 'duplicate_parameter' || responseData.message?.includes('already exist')) {
            console.log(`Contact exists, updating: ${contact.email}`);
            const updateResult = updateContactList(contact.email, apiKey);
            if (updateResult.success) {
              results.push({ email: contact.email, success: true, note: 'Updated existing contact' });
            } else {
              console.error(`Failed to update existing contact ${contact.email}:`, updateResult.error);
              results.push({ email: contact.email, success: false, error: updateResult.error });
            }
          } else {
            console.error(`Failed to import contact ${contact.email}:`, responseData);
            results.push({ email: contact.email, success: false, error: responseData });
          }
        } else {
          // Other error codes
          let responseData = '';
          try {
            const responseText = response.getContentText();
            if (responseText) {
              responseData = JSON.parse(responseText);
            }
          } catch (parseError) {
            responseData = `Response code: ${responseCode}`;
          }
          console.error(`Failed to import contact ${contact.email}:`, responseData);
          results.push({ email: contact.email, success: false, error: responseData });
        }
        
        // Small delay between API calls to respect rate limits
        Utilities.sleep(100); // 100ms delay
        
      } catch (error) {
        console.error(`Error importing contact ${contact.email}:`, error);
        results.push({ email: contact.email, success: false, error: error.toString() });
      }
    });
    
    // Consider batch successful if at least one contact was imported
    const successfulImports = results.filter(r => r.success);
    
    return {
      success: successfulImports.length > 0,
      results: results,
      error: successfulImports.length === 0 ? 'No contacts were successfully imported' : null
    };
    
  } catch (error) {
    console.error('Error in importContactsToBrevo:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Update an existing contact to add them to the list
 */
function updateContactList(email, apiKey) {
  try {
    const updateUrl = `${BREVO_API_ENDPOINT}/${encodeURIComponent(email)}`;
    
    const updatePayload = {
      listIds: [BREVO_LIST_ID] // Add to list
    };
    
    const response = UrlFetchApp.fetch(updateUrl, {
      method: 'PUT',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'api-key': apiKey
      },
      payload: JSON.stringify(updatePayload),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode === 204) {
      console.log(`Successfully updated contact to add to list: ${email}`);
      return { success: true };
    } else {
      let responseData = '';
      try {
        const responseText = response.getContentText();
        if (responseText) {
          responseData = JSON.parse(responseText);
        }
      } catch (parseError) {
        responseData = `Response code: ${responseCode}`;
      }
      return { success: false, error: responseData };
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Setup and return Gmail labels
 */
function setupLabels() {
  const labels = {};
  
  // Get or create import label
  labels.importLabel = GmailApp.getUserLabelByName(IMPORT_LABEL);
  if (!labels.importLabel) {
    labels.importLabel = GmailApp.createLabel(IMPORT_LABEL);
  }
  
  // Get or create success label
  labels.successLabel = GmailApp.getUserLabelByName(SUCCESS_LABEL);
  if (!labels.successLabel) {
    labels.successLabel = GmailApp.createLabel(SUCCESS_LABEL);
  }
  
  // Get or create error label
  labels.errorLabel = GmailApp.getUserLabelByName(ERROR_LABEL);
  if (!labels.errorLabel) {
    labels.errorLabel = GmailApp.createLabel(ERROR_LABEL);
  }
  
  return labels;
}

/**
 * Get Brevo API key from secure storage
 */
function getBrevoApiKey() {
  return PropertiesService.getScriptProperties().getProperty(BREVO_API_KEY_PROPERTY);
}

/**
 * Setup function to securely store the Brevo API key
 * Run this once to configure the API key
 * IMPORTANT: Delete TEMP_API_KEY from top of file after running this!
 */
function setupApiKey() {
  if (!TEMP_API_KEY) {
    throw new Error('TEMP_API_KEY not found. Please add it to the top of the file temporarily.');
  }
  
  PropertiesService.getScriptProperties().setProperty(BREVO_API_KEY_PROPERTY, TEMP_API_KEY);
  console.log('Brevo API key has been stored securely. NOW DELETE THE TEMP_API_KEY LINE FROM THE TOP OF THIS FILE!');
}

/**
 * Create a time-driven trigger to run the import process automatically
 * Runs every hour during business hours (9 AM - 5 PM)
 */
function createHourlyTrigger() {
  // Delete existing triggers first
  deleteAllTriggers();
  
  // Create new trigger
  ScriptApp.newTrigger('processBrevoImports')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log('Hourly trigger created for processBrevoImports');
}

/**
 * Create a time-driven trigger to run the import process daily
 * Runs once per day at 9 AM
 */
function createDailyTrigger() {
  // Delete existing triggers first
  deleteAllTriggers();
  
  // Create new trigger
  ScriptApp.newTrigger('processBrevoImports')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  console.log('Daily trigger created for processBrevoImports at 9 AM');
}

/**
 * Delete all existing triggers
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  console.log(`Deleted ${triggers.length} existing trigger(s)`);
}

/**
 * Test function to validate email extraction
 */
function testEmailExtraction() {
  const testText = `
    Please contact john.doe@example.com for more information.
    You can also reach out to "Jane Smith" <jane.smith@company.org>
    or Mary Johnson <mary@test.co.uk> for support.
    Another contact is support@help.net
  `;
  
  const emails = extractEmailAddresses(testText);
  console.log('Extracted emails:', emails);
}

/**
 * Test function to validate Brevo API connection
 */
function testBrevoConnection() {
  const testContact = [{ email: 'test@example.com', name: 'Test User' }];
  const result = importContactsToBrevo(testContact);
  console.log('Test import result:', result);
}

/**
 * Manual cleanup function to remove labels from all emails
 * Use with caution - this will remove labels from ALL emails
 */
function cleanupLabels() {
  const labelNames = [SUCCESS_LABEL, ERROR_LABEL, IMPORT_LABEL];
  
  labelNames.forEach(labelName => {
    const label = GmailApp.getUserLabelByName(labelName);
    if (label) {
      const threads = label.getThreads();
      console.log(`Removing ${labelName} from ${threads.length} thread(s)`);
      label.removeFromThreads(threads);
    }
  });
}