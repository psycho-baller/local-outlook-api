# Email Sender Application

This Node.js application reads recipient data from a CSV file and sends personalized emails by replacing placeholders in an email template. It uses the local Outlook API to send emails through the Outlook Web Client.

## Features

- Reads recipient data from a CSV file
- Uses a customizable HTML email template with placeholders
- Replaces placeholders in both subject line and email body
- Sends personalized emails to each recipient through the Outlook Web Client
- Integrates with the local API server that controls Outlook

## Setup

1. Install dependencies:
   ```
   npm install
   ```

## To run a mail batch

1. Customize the email template:
   - Edit `body.html` to change the email content and design
   - Placeholders are defined as `{{fieldName}}` where `fieldName` corresponds to a column in your CSV file

   Customize index.js - set EMAIL_SUBJECT - it contains the subject line of the email with placeholders

2. Prepare your recipient data:
   - Update `recipients.csv` with your recipient data
   - The first column must be `emailAddress`
   - Add any additional columns needed for your template placeholders

3. Start the local API server:
   - Navigate to the `webFormAndSSE` directory and run:
   ```
   npm start
   ```
   - Make sure the server is running on http://localhost:3000
   - Keep this terminal window open

4. Ensure you're logged into Outlook Web Client in your browser
   - The local API server communicates with the Outlook Web Client through browser automation
   - Keep this browser window open
   - Open Browser Manage Extensions window; edit options for the Outlook Email Assistant
   - Set the SSE Endpoint Url to http://localhost:3000/events and click Connect; the log should show "Connected to SSE endpoint"; when it does, the extension is listening for SSE events that instruct it to send emails


## Usage

1. Make sure the local API server is running (see Step 3)

2. Run the email sender application:

```
npm start
```

3. The application will:
   - Check if the server is running
   - Read the CSV file and email template
   - Process each recipient and send personalized emails
   - Log the results to the console

## CSV File Format

The application supports both comma-separated and semicolon-separated CSV files, with or without double quotes around field values.

The CSV file should have the following format:
- First row: Header with column names
- The file should contain an email column (can be named `email`, `emailAddress`, `Email`, or `e-mail`)
- Additional columns: Can be named anything and used as placeholders in the template

### Examples

**Comma-separated format:**
```
emailAddress,name,product,duration,plan
john.doe@example.com,John Doe,Premium Package,6 months,Business Pro
```

**Semicolon-separated format with quotes:**
```
"Bedrijf";"Coordinator";"Naam";"email"
"theFactor.e";"";"René Kamp";"r.kamp@tfe.nl"
```

**Semicolon-separated format without quotes:**
```
Bedrijf;Coordinator;Naam;email
theFactor.e;;René Kamp;r.kamp@tfe.nl
```

## Email Template

The email template uses Handlebars syntax for placeholders. Any placeholder in the format `{{fieldName}}` will be replaced with the corresponding value from the CSV file.

## Subject Line

The subject line is defined as a constant in `index.js`. It also supports placeholders in the same format as the email body.


## Tab to CSV

You can use the `tab-to-csv.js` script to convert a tab-separated file to a CSV file. (that is convenient because copy paste from Excel results in tab-separated files)

I've created a Node.js script that converts tab-delimited files to semicolon-separated files with double quotes around field values. The script is saved as tab-to-csv.js in your email-sender directory.

Features of the Script
- Converts tab-delimited files to semicolon-separated format
- Adds double quotes around all field values
- Properly escapes any existing double quotes in the data
- Trims whitespace from field values
- Provides detailed console output during conversion

### How to Use the Script

Run the script with Node.js, providing the input file path as an argument:

```bash
node tab-to-csv.js <input-file> [output-file]
```

If you don't specify an output file, it will automatically create one with the same name as the input file but with a .csv extension.

### Example

```bash
node tab-to-csv.js data.txt
```

This will convert `data.txt` (tab-delimited) to `data.csv` (semicolon-separated with quoted values).

Or specify a custom output file:

```bash
node tab-to-csv.js data.txt converted-data.csv
```

Sample Conversion
Input (tab-delimited):

```text
Name	Age	City
John Doe	30	New York
Jane Smith	25	Los Angeles
```

Output (semicolon-separated with quoted values):

```text
"Name";"Age";"City"
"John Doe";"30";"New York"
"Jane Smith";"25";"Los Angeles"
```

The script handles all the necessary formatting and escaping to ensure your data is properly converted.