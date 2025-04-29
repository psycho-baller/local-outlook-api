/**
 * Outlook Email Extension - Programmatic Email Client
 * 
 * This Node.js application sends emails programmatically through the SSE server
 * using the same endpoint as the HTML form client.
 */

const http = require('http');
const readline = require('readline');

// Configuration
const config = {
  serverUrl: 'http://localhost:3000',
  defaultEmail: '',
  defaultSubject: 'Automated email from Node.js client',
  defaultBody: '<p>This email was sent programmatically through the Outlook Email Extension.</p>'
};

// Create readline interface for user input
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

/**
 * Sends an email through the SSE server
 * @param {Object} emailData - Email data (to, subject, body)
 * @returns {Promise<Object>} - Response from the server
 */
async function sendEmail(emailData) {
  return new Promise((resolve, reject) => {
    console.log('\nSending email...');
    
    // Prepare the request data
    const postData = JSON.stringify(emailData);
    
    // Parse the server URL
    const serverUrl = new URL(config.serverUrl);
    
    // Request options
    const options = {
      hostname: serverUrl.hostname,
      port: serverUrl.port,
      path: '/send-email',
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
          console.log('\n✅ Request submitted successfully!');
          console.log('The server is processing your request. Check the server logs for updates.');
          resolve({ success: true });
        } else {
          console.error(`\n❌ Failed to submit email request: ${res.statusCode} ${res.statusMessage}`);
          resolve({ success: false, error: res.statusMessage });
        }
      });
    });
    
    // Handle request errors
    req.on('error', (error) => {
      console.error('\n❌ Error sending email:', error.message);
      resolve({ success: false, error: error.message });
    });
    
    // Send the request data
    req.write(postData);
    req.end();
  });
}

/**
 * Prompts the user for email details
 * @returns {Promise<Object>} - Email data
 */
async function promptEmailDetails() {
  return new Promise((resolve) => {
    console.log('\n📧 Outlook Email Extension - Programmatic Client');
    console.log('=============================================');
    
    rl.question(`Recipient email (${config.defaultEmail ? 'default: ' + config.defaultEmail : 'required'}): `, (to) => {
      to = to.trim() || config.defaultEmail;
      
      if (!to) {
        console.error('\n❌ Recipient email is required!');
        return promptEmailDetails().then(resolve);
      }
      
      rl.question(`Subject (default: ${config.defaultSubject}): `, (subject) => {
        subject = subject.trim() || config.defaultSubject;
        
        console.log('\nEmail body (HTML supported, enter a single dot "." on a new line to finish):');
        let body = '';
        let bodyInput = '';
        
        rl.on('line', (line) => {
          if (line.trim() === '.') {
            // Use provided body or default
            body = body.trim() || config.defaultBody;
            
            // Resolve with email data
            resolve({ to, subject, body });
            
            // Remove the line listener to prevent memory leaks
            rl.removeAllListeners('line');
          } else {
            bodyInput += line + '\n';
            body = bodyInput;
          }
        });
      });
    });
  });
}

/**
 * Main function
 */
async function main() {
  try {
    // Check if server is running
    try {
      await new Promise((resolve, reject) => {
        const serverUrl = new URL(config.serverUrl);
        const req = http.get({
          hostname: serverUrl.hostname,
          port: serverUrl.port,
          path: '/',
          method: 'GET'
        }, (res) => {
          if (res.statusCode === 200) {
            resolve();
          } else {
            reject(new Error(`Server returned status code ${res.statusCode}`));
          }
        });
        
        req.on('error', reject);
        req.end();
      });
      
      console.log(`✅ Server is running at ${config.serverUrl}`);
    } catch (error) {
      console.error(`❌ Cannot connect to server at ${config.serverUrl}`);
      console.error('Please make sure the server is running.');
      process.exit(1);
    }
    
    // Get email details from user
    const emailData = await promptEmailDetails();
    
    // Send the email
    await sendEmail(emailData);
    
    // Ask if user wants to send another email
    rl.question('\nDo you want to send another email? (y/n): ', (answer) => {
      if (answer.toLowerCase() === 'y') {
        main();
      } else {
        console.log('\nThank you for using the Outlook Email Extension client!');
        rl.close();
      }
    });
  } catch (error) {
    console.error('An unexpected error occurred:', error);
    rl.close();
  }
}

// Start the application
main();
