# Outlook Email SSE Server

This is a NodeJS application that implements a Server-Sent Events (SSE) endpoint for the Outlook Email Extension. It also provides a web form for sending email instructions to the extension.

## Features

- SSE endpoint for sending real-time email instructions to the Outlook Email Extension
- Web form for composing and sending email instructions
- Real-time connection status indicator
- HTML preview for email body content

## Prerequisites

- Node.js (v14 or higher)
- npm (v6 or higher)

## Installation

1. Navigate to the project directory:
   ```
   cd webFormAndSSE
   ```

2. Install the dependencies:
   ```
   npm install
   ```

## Usage

1. Start the server:
   ```
   npm start
   ```

2. The server will be running at:
   - Web form: http://localhost:3000
   - SSE endpoint: http://localhost:3000/events

3. Connect the Outlook Email Extension to the SSE endpoint:
   - Open the extension settings (right-click the extension icon → Options)
   - Enter the SSE endpoint URL: `http://localhost:3000/events`
   - Click "Connect"

4. Use the web form to send email instructions:
   - Fill in the recipient, subject, and body fields
   - Click "Send Email Instruction"
   - The instruction will be sent to all connected extensions via SSE

## Development

For development with auto-restart on file changes:
```
npm run dev
```

## API

### SSE Endpoint

- **URL**: `/events`
- **Method**: `GET`
- **Description**: Establishes a Server-Sent Events connection for sending email instructions to the Outlook Email Extension.

### Send Email Instruction

- **URL**: `/send-email`
- **Method**: `POST`
- **Form Parameters**:
  - `to`: Recipient email address
  - `subject`: Email subject
  - `body`: Email body content (HTML supported)
- **Description**: Sends an email instruction to all connected SSE clients.

## Event Types

The server sends the following event types:

- `connection`: Sent when a client connects to the SSE endpoint
- `email-instruction`: Sent when an email instruction is submitted via the web form
