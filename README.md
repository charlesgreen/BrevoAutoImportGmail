# Brevo Auto Import Gmail

A Google Apps Script that automatically imports email addresses from Gmail messages into Brevo CRM. The script monitors emails with a specific label, extracts contact information, and imports them to Brevo using their API.

## Features

- **Automated Email Processing**: Monitors Gmail for emails labeled `Brevo/Import`
- **Smart Email Extraction**: Uses multiple regex patterns to extract email addresses and names from email bodies
- **Brevo CRM Integration**: Imports contacts using Brevo API v3 with proper error handling
- **Label Management**: Automatically applies success (`Brevo/Success`) or error (`Brevo/Error`) labels
- **Batch Processing**: Processes emails in batches to avoid timeouts
- **Rate Limiting**: Includes delays to respect API rate limits
- **Secure Configuration**: Stores API keys securely using PropertiesService
- **Comprehensive Logging**: Detailed console logging for monitoring and debugging

## Setup Instructions

### 1. Create Google Apps Script Project

1. Go to [Google Apps Script](https://script.google.com)
2. Click "New Project"
3. Replace the default `Code.gs` content with the provided `Code.gs` file
4. Add the `appsscript.json` manifest file to configure permissions

### 2. Configure API Key and List ID

#### API Key Setup

1. Add your Brevo API key to the top of `Code.gs` temporarily:
   ```javascript
   const TEMP_API_KEY = 'your-brevo-api-key-here';
   ```

2. Run the `setupApiKey()` function once to store it securely:
   ```javascript
   setupApiKey(); // Stores the API key securely
   ```

3. **IMPORTANT**: Delete the `TEMP_API_KEY` line from the code after running `setupApiKey()`

#### List Configuration

By default, contacts are added to list #14. To change this, modify the `BREVO_LIST_ID` constant:

```javascript
const BREVO_LIST_ID = 14; // Change this to your desired list ID
```

### 3. Set Up Gmail Labels

The script will automatically create these labels if they don't exist:

- `Brevo/Import` - Apply this label to emails containing contacts to import
- `Brevo/Success` - Applied to emails after successful import
- `Brevo/Error` - Applied to emails that failed to import

### 4. Configure Triggers

Choose one of the following trigger options:

#### Option A: Hourly Processing

```javascript
createHourlyTrigger(); // Runs every hour
```

#### Option B: Daily Processing

```javascript
createDailyTrigger(); // Runs once daily at 9 AM
```

## Usage

### Manual Processing

Run `processBrevoImports()` manually to process all emails with the `Brevo/Import` label.

### Automatic Processing

Set up triggers using the provided functions to run automatically.

### Email Format Support

The script can extract email addresses from various formats:

```text
Simple format: john@example.com
With display name: "John Doe" <john@example.com>
Without quotes: John Doe <john@example.com>
Multiple emails in one message
```

## API Integration

### Brevo API Details

- **Endpoint**: `https://api.brevo.com/v3/contacts`
- **Method**: POST
- **Authentication**: API key in header (`api-key`)
- **Rate Limiting**: 100ms delay between requests

### Contact Data Structure

```json
{
  "email": "contact@example.com",
  "attributes": {
    "FIRSTNAME": "John",
    "LASTNAME": "Doe"
  },
  "listIds": [14],
  "updateEnabled": true
}
```

## Functions Reference

### Core Functions

- `processBrevoImports()` - Main processing function
- `setupApiKey()` - Configure API key (run once)
- `setupLabels()` - Create Gmail labels
- `extractEmailAddresses(text)` - Extract emails from text
- `importContactsToBrevo(contacts)` - Import contacts to Brevo
- `updateContactList(email, apiKey)` - Update existing contact to add to list

### Trigger Management

- `createHourlyTrigger()` - Set up hourly processing
- `createDailyTrigger()` - Set up daily processing
- `deleteAllTriggers()` - Remove all triggers

### Testing Functions

- `testEmailExtraction()` - Test email extraction patterns
- `testBrevoConnection()` - Test Brevo API connection
- `cleanupLabels()` - Remove labels from all emails (use with caution)

## Error Handling

The script includes comprehensive error handling:

- **API Errors**: Logged with response details
- **Duplicate Contacts**: Automatically updates existing contacts to add them to the list
- **Network Errors**: Caught and logged with retry logic
- **Parsing Errors**: Individual contact failures don't stop batch processing
- **Timeout Protection**: Maximum execution time of 4 minutes

## Monitoring and Logging

Check the Google Apps Script logs for:

- Processing statistics
- Individual contact import results
- Error details and troubleshooting information
- API response codes and messages

## Security Considerations

- API keys are stored using PropertiesService (encrypted)
- No sensitive data is logged
- API requests use HTTPS
- OAuth scopes are limited to necessary permissions

## Required Permissions

The script requires these OAuth scopes:

- `gmail.modify` - Read and modify Gmail messages and labels
- `gmail.labels` - Create and manage Gmail labels
- `script.external_request` - Make HTTP requests to Brevo API

## Troubleshooting

### Common Issues

1. **No emails processed**: Check if emails have the `Brevo/Import` label
2. **API errors**: Verify API key is correctly configured
3. **Permission errors**: Ensure all OAuth scopes are authorized
4. **Timeout errors**: Reduce batch size or increase delays

### Debugging Steps

1. Run `testEmailExtraction()` to verify regex patterns
2. Run `testBrevoConnection()` to check API connectivity
3. Check Google Apps Script logs for detailed error messages
4. Verify Gmail labels exist and are correctly applied

## Limitations

- Maximum 6 minutes execution time per run (Google Apps Script limit)
- Processes up to 50 email threads per run
- API rate limits may affect processing speed
- Requires manual labeling of emails for import

## Contributing

This script is designed to be self-contained and production-ready. Modify the constants at the top of the file to customize behavior:

- `BREVO_LIST_ID`: Target list ID for imported contacts (default: 14)
- `BATCH_SIZE`: Number of emails processed per batch
- `MAX_PROCESSING_TIME`: Maximum execution time in milliseconds
- Label names and API endpoints

## License

See LICENSE file for license information.
