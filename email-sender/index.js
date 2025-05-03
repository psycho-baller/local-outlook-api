const fs = require('fs-extra');
const csv = require('csv-parser');
const handlebars = require('handlebars');
const path = require('path');
const http = require('http');

// Constants
const CSV_FILE_PATH = path.join(__dirname, 'recipients.csv');
const EMAIL_TEMPLATE_PATH = path.join(__dirname, 'body.html');

// Email subject - you may use any placeholder defined in the CSV file
const EMAIL_SUBJECT = 'Pilot AI Assisted Coding @ {{bedrijf}} ';

// Server configuration
const SERVER_CONFIG = {
  serverUrl: 'http://localhost:3000',
  endpoint: '/send-email'
};

// Ensure the server is running before starting the email sending process
async function checkServerConnection() {
  return new Promise((resolve, reject) => {
    const serverUrl = new URL(SERVER_CONFIG.serverUrl);
    const req = http.get({
      hostname: serverUrl.hostname,
      port: serverUrl.port,
      path: '/',
      method: 'GET'
    }, (res) => {
      if (res.statusCode === 200) {
        console.log(`✅ Server is running at ${SERVER_CONFIG.serverUrl}`);
        resolve(true);
      } else {
        reject(new Error(`Server returned status code ${res.statusCode}`));
      }
    });
    
    req.on('error', (error) => {
      console.error(`❌ Cannot connect to server at ${SERVER_CONFIG.serverUrl}`);
      console.error('Please make sure the server is running.');
      reject(error);
    });
    
    req.end();
  });
}

/**
 * Replace placeholders in a template string using data from a record
 * @param {string} template - Template string with {{placeholders}}
 * @param {Object} data - Data object with values to replace placeholders
 * @returns {string} - Processed string with placeholders replaced
 */
function processTemplate(template, data) {
  const compiledTemplate = handlebars.compile(template);
  return compiledTemplate(data);
}

/**
 * Send an email using the local API that wraps the Outlook client
 * @param {Object} emailData - Email data including recipient, subject, and body
 * @returns {Promise} - Promise that resolves when email is sent
 */
async function sendEmail(emailData) {
  return new Promise((resolve, reject) => {
    console.log(`Sending email to ${emailData.to}...`);
    
    // Prepare the request data
    const postData = JSON.stringify({
      to: emailData.to,
      subject: emailData.subject,
      body: emailData.body
    });
    
    // Parse the server URL
    const serverUrl = new URL(SERVER_CONFIG.serverUrl);
    
    // Request options
    const options = {
      hostname: serverUrl.hostname,
      port: serverUrl.port,
      path: SERVER_CONFIG.endpoint,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(postData)
      }
    };
    
    // Make the request
    const req = http.request(options, (res) => {
      let responseData = '';
      
      // Collect response data
      res.on('data', (chunk) => {
        responseData += chunk;
      });
      
      // Process the complete response
      res.on('end', () => {
        if (res.statusCode === 200) {
          console.log(`✅ Email to ${emailData.to} submitted successfully!`);
          resolve({ success: true });
        } else {
          console.error(`❌ Failed to submit email to ${emailData.to}: ${res.statusCode} ${res.statusMessage}`);
          resolve({ success: false, error: res.statusMessage });
        }
      });
    });
    
    // Handle request errors
    req.on('error', (error) => {
      console.error(`❌ Error sending email to ${emailData.to}:`, error.message);
      resolve({ success: false, error: error.message });
    });
    
    // Send the request data
    req.write(postData);
    req.end();
  });
}

/**
 * Main function to process CSV and send emails
 */
async function main() {
  try {
    // Check if the server is running
    await checkServerConnection();
    
    // Read email template
    const emailTemplateContent = await fs.readFile(EMAIL_TEMPLATE_PATH, 'utf8');
    
    // Process CSV file
    const records = [];
    
    // Read and parse CSV file
    await new Promise((resolve, reject) => {
      fs.createReadStream(CSV_FILE_PATH)
        .pipe(csv({
          separator: ';',  // Use semicolon as separator
          quote: '"',     // Use double quotes for quoted fields
          escape: '"'     // Use double quotes for escaping
        }))
        .on('data', (data) => records.push(data))
        .on('end', resolve)
        .on('error', reject);
    });
    
    console.log(`Found ${records.length} recipients in CSV file`);
    // show the field names 
    console.log(records[0]);
    // Process each record and send email
    for (const record of records) {
      try {
        // Process email subject with placeholders
        const subject = processTemplate(EMAIL_SUBJECT, record);
        
        // Process email body with placeholders
        const body = processTemplate(emailTemplateContent, record);
        
        // Determine which field contains the email address
        const emailField = record.email || record.emailAddress || record.Email || record['e-mail'];
        
        if (!emailField) {
          console.error(`❌ No email address found for record:`, record);
          continue;
        }
        
        // Prepare email data
        const emailData = {
          to: emailField,
          subject,
          body
        };
        
        // Send email
        const result = await sendEmail(emailData);
        
        if (result.success) {
          console.log(`✅ Successfully processed email for: ${emailData.to}`);
        } else {
          console.error(`❌ Failed to send email to ${emailData.to}: ${result.error || 'Unknown error'}`);
        }
        
        // Add a small delay between requests to avoid overwhelming the server
        await new Promise(resolve => setTimeout(resolve, 1000));
      } catch (error) {
        const emailField = record.email || record.emailAddress || record.Email || record['e-mail'] || 'unknown';
        console.error(`Error processing record for ${emailField}:`, error);
      }
    }
    
    console.log('Email sending process completed');
  } catch (error) {
    console.error('Error in main process:', error);
  }
}

// Run the main function
main().catch(console.error);
