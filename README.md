# Outlook Email Assistant Chrome Extension

This Chrome extension integrates with Outlook Web Client to send emails and retrieve event details through a convenient side panel interface. It also supports receiving instructions via Server-Sent Events (SSE) from a server endpoint, with callback functionality to report operation results.

## Features

- Side panel interface for composing emails and retrieving event details
- Works with Outlook Web Client (Office 365 and Outlook.com)
- Supports HTML content in email body
- Automatically saves draft emails and previous searches
- Server-Sent Events (SSE) integration for receiving instructions from a server
- Callback functionality to report operation results back to the server
- Automatic tab management (finds or creates Outlook tabs as needed)
- Event search across multiple months
- Extraction of event attendees and details
- Programmatic email client for Node.js applications

## Installation

1. Open Chrome and navigate to `chrome://extensions/`
2. Enable "Developer mode" by toggling the switch in the top-right corner
3. Click "Load unpacked" and select the `outlook-email-extension` folder
4. The extension should now appear in your Chrome toolbar

## Usage

### Manual Email Sending

1. Navigate to [Outlook Web Client](https://outlook.office.com) or [Outlook.com](https://outlook.live.com)
2. Click the extension icon in your Chrome toolbar
3. The side panel will open with the "Compose Email" tab active
4. Compose your email:
   - Enter recipient email address
   - Enter subject
   - Enter body content (HTML is supported)
   - Click "Send Email"

### Event Details Retrieval

1. Navigate to [Outlook Web Client](https://outlook.office.com) or [Outlook.com](https://outlook.live.com)
2. Click the extension icon in your Chrome toolbar
3. Switch to the "Event Details" tab in the side panel
4. Enter an event title or ID (partial titles work)
5. Click "Get Event Details"
6. View the event information including:
   - Event title, organizer, date/time, and location
   - List of attendees with their names and email addresses
7. Use the "Copy All Emails" button to copy attendee emails to clipboard

### Server-Sent Events (SSE) Integration

1. Open the extension settings (right-click the extension icon → Options)
2. Enter your SSE endpoint URL (e.g., `https://your-server.com/events`)
3. Click "Connect"
4. The extension will now listen for instructions from your server

#### Email Instructions

Send events from your server with the type `email-instruction` and data in this format:
```json
{
  "to": "recipient@example.com",
  "subject": "Email Subject",
  "body": "<p>Email body content (HTML supported)</p>",
  "requestId": "unique-request-id",
  "callbackUrl": "https://your-server.com/email-result"
}
```

#### Event Attendee Requests

Send events with the type `get-event-attendees` and data in this format:
```json
{
  "eventId": "Meeting Title or ID",
  "requestId": "unique-request-id",
  "callbackUrl": "https://your-server.com/event-attendees-result"
}
```

#### Callback Functionality

The extension will send the result of operations back to the provided callback URL with:
- The original `requestId` to correlate with the request
- Success/failure status
- Any error messages if applicable
- Operation-specific result data (e.g., attendee list)

## Development

### Directory Structure

```
outlook-email-extension/
├── manifest.json         # Extension configuration
├── src/                  # Source code directory
│   ├── background/       # Background scripts
│   │   └── background.js # Background service worker
│   ├── content/          # Content scripts
│   │   ├── content.js    # Main content script (message router)
│   │   ├── email.js      # Email-related functionality
│   │   ├── calendar.js   # Calendar and event-related functionality
│   │   └── utils.js      # Shared utility functions
│   ├── ui/               # User interface components
│   │   ├── panel/        # Side panel components
│   │   │   ├── panel.html# Side panel UI
│   │   │   └── panel.js  # Side panel logic
│   │   ├── popup/        # Popup components
│   │   │   ├── popup.html# Popup UI
│   │   │   └── popup.js  # Popup logic
│   │   └── settings/     # Settings components
│   │       ├── settings.html# Settings UI
│   │       └── settings.js  # Settings logic
│   └── styles/           # Stylesheets
│       └── styles.css    # Styles for UI components
├── icons/                # Extension icons
└── webFormAndSSE/        # Server components
    ├── server.js         # NodeJS SSE server
    ├── views/            # EJS templates for web interface
    ├── public/           # Static assets
    └── email-client.js   # Programmatic email client
```

### How It Works

#### Modular Content Script Architecture
1. The main content script (content.js) serves as a message router
2. Specialized scripts handle specific functionality:
   - email.js: Handles all email-related operations
   - calendar.js: Manages calendar and event-related functionality
   - utils.js: Provides shared utility functions
3. Scripts are loaded in the correct order via the manifest.json file
4. This modular approach improves maintainability and separation of concerns while maintaining compatibility with Chrome Extension requirements

#### Manual Email Sending
1. The extension checks if you're on an Outlook Web Client page
2. If you are, it enables the side panel for composing emails
3. When you submit the form, the email.js module interacts with the Outlook interface to create and send the email

#### Event Details Retrieval
1. When you enter an event title/ID and click "Get Event Details"
2. The calendar.js module searches for the event in the current month view
3. If not found, it checks the next two months automatically
4. When found, it extracts event details and attendee information
5. The results are displayed in the side panel

#### SSE Integration
1. The background script establishes and maintains a connection to your SSE endpoint
2. When an event is received, it validates the data and processes the instruction
3. It finds or creates an Outlook tab as needed
4. The appropriate content script module interacts with the Outlook interface to perform the requested action
5. Results are sent back to the server via the provided callback URL

#### Programmatic Email Client
1. A Node.js application that uses the same endpoint as the web form
2. Allows sending emails programmatically without the web interface
3. Supports interactive command-line input or can be integrated into other applications

## Notes

- The extension requires permission to access Outlook Web Client pages and your SSE endpoint
- It does not collect or transmit any data outside of the Outlook interface and your configured SSE endpoint
- Draft emails and previous event searches are saved locally in your browser storage
- The SSE connection will attempt to reconnect automatically if disconnected
- You can test the SSE functionality using the test email feature in the settings page
- The event search functionality can find events by partial title match
- Event search automatically checks future months if the event isn't found in the current month
- The programmatic email client requires Node.js but has minimal dependencies
- Content scripts are organized in a modular fashion but don't use ES modules due to Chrome Extension limitations

## Programmatic Email Client

### Installation

```bash
cd webFormAndSSE
```

### Usage

```bash
node email-client.js
```

Follow the interactive prompts to send an email programmatically:

1. Enter recipient email address (or use the default)
2. Enter subject (or use the default)
3. Type the email body (HTML is supported)
4. Enter a single dot (`.`) on a new line to finish the body

The client will connect to the SSE server and send the email instruction, then wait for the result.

### Integration

The email client can be modified to work as a module in other Node.js applications:

```javascript
const { sendEmail } = require('./email-client');

sendEmail({
  to: 'recipient@example.com',
  subject: 'Automated Email',
  body: '<p>Hello from my application!</p>'
}).then(result => {
  console.log('Email sent:', result);
});
```
