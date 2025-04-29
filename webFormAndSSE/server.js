const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Set up middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Set up EJS as the view engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Store connected SSE clients
const clients = [];

// Store pending email instruction requests
const pendingEmailRequests = new Map();

// Create a dedicated route for SSE
app.get('/events', function(req, res) {
  // Log connection attempt
  const timestamp = new Date().toISOString();
  console.log(`[${timestamp}] SSE connection attempt from ${req.ip}`);

  // Set required headers for SSE
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', '*');
  
  // Ensure the response is sent immediately
  res.flushHeaders();
  
  // Send an initial comment to establish the connection
  res.write(':\n\n');
  
  // Create client object
  const clientId = Date.now();
  const newClient = {
    id: clientId,
    response: res
  };
  
  // Add this client to the connected clients
  clients.push(newClient);
  
  // Enhanced connection logging
  const clientIP = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
  console.log(`[${timestamp}] Client connected - ID: ${clientId}, IP: ${clientIP}, User-Agent: ${req.headers['user-agent']}`);
  console.log(`[${timestamp}] Total connected clients: ${clients.length}`);
  
  // Send a welcome message
  sendEventToClient(newClient, 'connection', { message: 'Connected to SSE server' });
  
  // Handle client disconnect
  req.on('close', () => {
    // Enhanced disconnection logging
    const disconnectTime = new Date().toISOString();
    console.log(`[${disconnectTime}] Client disconnected - ID: ${clientId}, IP: ${clientIP}`);
    
    // Remove this client from the array
    const index = clients.findIndex(client => client.id === clientId);
    if (index !== -1) {
      clients.splice(index, 1);
      console.log(`[${disconnectTime}] Client removed from active connections. Total remaining: ${clients.length}`);
    } else {
      console.log(`[${disconnectTime}] Warning: Could not find client ID ${clientId} in active connections`);
    }
  });
});

// Home page - render the email form
app.get('/', (req, res) => {
  res.render('index', { 
    title: 'Outlook Email Assistant',
    connected: clients.length,
    success: null,
    page: 'email'
  });
});

// Event attendees page
app.get('/event-attendees', (req, res) => {
  res.render('event-attendees', { 
    title: 'Event Attendees Retriever',
    connected: clients.length,
    success: null,
    page: 'attendees'
  });
});

// Handle email form submission
app.post('/send-email', (req, res) => {
  const { to, subject, body } = req.body;
  
  // Validate input
  if (!to || !subject || !body) {
    return res.render('index', { 
      title: 'Outlook Email Assistant',
      connected: clients.length,
      error: 'All fields are required',
      formData: req.body,
      page: 'email'
    });
  }
  
  // Create a unique request ID
  const requestId = `email_${Date.now()}`;
  
  // Create email instruction with request ID and callback URL
  const emailInstruction = {
    to,
    subject,
    body,
    requestId,
    callbackUrl: `${req.protocol}://${req.get('host')}/email-result`
  };
  
  // Store the request in pending requests
  pendingEmailRequests.set(requestId, {
    timestamp: Date.now(),
    request: emailInstruction,
    res: res // Store the response object to respond later
  });
  
  // Set a timeout to clean up if no response is received
  setTimeout(() => {
    if (pendingEmailRequests.has(requestId)) {
      const pendingRequest = pendingEmailRequests.get(requestId);
      pendingEmailRequests.delete(requestId);
      
      // Only send timeout response if the response hasn't been sent yet
      if (!pendingRequest.res.headersSent) {
        pendingRequest.res.render('index', { 
          title: 'Outlook Email Assistant',
          connected: clients.length,
          error: 'Request timed out. Please try again.',
          formData: req.body,
          page: 'email'
        });
      }
    }
  }, 300000); // 5 minute timeout 
  
  // Send to all connected clients
  const clientCount = sendEventToAllClients('email-instruction', emailInstruction);
  
  if (clientCount === 0) {
    // Clean up the pending request
    pendingEmailRequests.delete(requestId);
    
    return res.render('index', { 
      title: 'Outlook Email Assistant',
      connected: clients.length,
      error: 'No connected clients to receive the email instruction',
      formData: req.body,
      page: 'email'
    });
  }
  
  // The response will be sent when the result is received
});

// Store pending event attendee requests
const pendingAttendeeRequests = new Map();

// Endpoint to receive email results
app.post('/email-result', (req, res) => {
  const { requestId, success, error } = req.body;
  
  // Log the received result
  console.log(`Received email result for request ${requestId}:`, req.body);
  
  // Check if this is a pending request
  if (pendingEmailRequests.has(requestId)) {
    const pendingRequest = pendingEmailRequests.get(requestId);
    pendingEmailRequests.delete(requestId);
    
    // Send the response to the waiting client
    if (!pendingRequest.res.headersSent) {
      if (success) {
        pendingRequest.res.render('index', { 
          title: 'Outlook Email Assistant',
          connected: clients.length,
          success: `Email sent successfully to ${pendingRequest.request.to}`,
          formData: {},
          page: 'email'
        });
      } else {
        pendingRequest.res.render('index', { 
          title: 'Outlook Email Assistant',
          connected: clients.length,
          error: error || 'Failed to send email',
          formData: { 
            to: pendingRequest.request.to,
            subject: pendingRequest.request.subject,
            body: pendingRequest.request.body
          },
          page: 'email'
        });
      }
    }
  }
  
  // Send a response to the extension
  res.json({ success: true });
});

// Handle event attendees form submission
app.post('/get-event-attendees', (req, res) => {
  const { eventId } = req.body;
  
  // Validate input
  if (!eventId) {
    return res.render('event-attendees', { 
      title: 'Event Attendees Retriever',
      connected: clients.length,
      error: 'Event ID is required',
      formData: req.body,
      page: 'attendees'
    });
  }
  
  // Create a unique request ID
  const requestId = `req_${Date.now()}`;
  
  // Create event attendees request
  const attendeesRequest = {
    eventId,
    requestId,
    callbackUrl: `${req.protocol}://${req.get('host')}/event-attendees-result`
  };
  
  // Store the request in pending requests
  pendingAttendeeRequests.set(requestId, {
    timestamp: Date.now(),
    request: attendeesRequest,
    res: res // Store the response object to respond later
  });
  
  // Set a timeout to clean up if no response is received
  setTimeout(() => {
    if (pendingAttendeeRequests.has(requestId)) {
      const pendingRequest = pendingAttendeeRequests.get(requestId);
      pendingAttendeeRequests.delete(requestId);
      
      // Only send timeout response if the response hasn't been sent yet
      if (!pendingRequest.res.headersSent) {
        pendingRequest.res.render('event-attendees', { 
          title: 'Event Attendees Retriever',
          connected: clients.length,
          error: 'Request timed out. Please try again.',
          formData: req.body,
          page: 'attendees'
        });
      }
    }
  }, 30000); // 30 second timeout
  
  // Send to all connected clients
  const clientCount = sendEventToAllClients('get-event-attendees', attendeesRequest);
  
  if (clientCount === 0) {
    // Clean up the pending request
    pendingAttendeeRequests.delete(requestId);
    
    return res.render('event-attendees', { 
      title: 'Event Attendees Retriever',
      connected: clients.length,
      error: 'No connected clients to process the request',
      formData: req.body,
      page: 'attendees'
    });
  }
  
  // The response will be sent when the result is received
});

// Endpoint to receive event attendees results
app.post('/event-attendees-result', (req, res) => {
  const { requestId, success, attendees, count, error } = req.body;
  
  // Log the received result
  console.log(`Received attendees result for request ${requestId}:`, req.body);
  
  // Check if this is a pending request
  if (pendingAttendeeRequests.has(requestId)) {
    const pendingRequest = pendingAttendeeRequests.get(requestId);
    pendingAttendeeRequests.delete(requestId);
    
    // Send the response to the waiting client
    if (!pendingRequest.res.headersSent) {
      if (success) {
        pendingRequest.res.render('event-attendees', { 
          title: 'Event Attendees Retriever',
          connected: clients.length,
          success: `Retrieved ${count} attendee(s) for event: ${pendingRequest.request.eventId}`,
          attendees: attendees,
          formData: {},
          page: 'attendees'
        });
      } else {
        pendingRequest.res.render('event-attendees', { 
          title: 'Event Attendees Retriever',
          connected: clients.length,
          error: error || 'Failed to retrieve attendees',
          formData: { eventId: pendingRequest.request.eventId },
          page: 'attendees'
        });
      }
    }
  }
  
  // Send a response to the extension
  res.json({ success: true });
});

// Function to send an event to a specific client
function sendEventToClient(client, eventType, data) {
  try {
    client.response.write(`event: ${eventType}\n`);
    client.response.write(`data: ${JSON.stringify(data)}\n\n`);
  } catch (error) {
    console.error(`Error sending event to client ${client.id}:`, error);
  }
}

// Function to send an event to all connected clients
function sendEventToAllClients(eventType, data) {
  let count = 0;
  const timestamp = new Date().toISOString();
  
  console.log(`[${timestamp}] Broadcasting event: ${eventType} to all clients`);
  
  clients.forEach(client => {
    sendEventToClient(client, eventType, data);
    count++;
  });
  
  console.log(`[${timestamp}] Event sent to ${count} client(s)`);
  return count;
}

// Add a test endpoint to verify SSE is working
app.get('/test-sse', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>SSE Test</title>
    </head>
    <body>
      <h1>SSE Test Client</h1>
      <div id="status">Connecting...</div>
      <div id="events"></div>
      
      <script>
        const eventsDiv = document.getElementById('events');
        const statusDiv = document.getElementById('status');
        
        const eventSource = new EventSource('/events');
        
        eventSource.onopen = () => {
          statusDiv.textContent = 'Connected to SSE';
          statusDiv.style.color = 'green';
        };
        
        eventSource.onerror = () => {
          statusDiv.textContent = 'Error connecting to SSE';
          statusDiv.style.color = 'red';
        };
        
        eventSource.addEventListener('connection', (event) => {
          const data = JSON.parse(event.data);
          const div = document.createElement('div');
          div.textContent = 'Connection event: ' + data.message;
          div.style.color = 'blue';
          eventsDiv.appendChild(div);
        });
        
        eventSource.addEventListener('email-instruction', (event) => {
          const data = JSON.parse(event.data);
          const div = document.createElement('div');
          div.textContent = 'Email instruction received: To: ' + data.to;
          div.style.color = 'green';
          eventsDiv.appendChild(div);
        });
        
        eventSource.onmessage = (event) => {
          const div = document.createElement('div');
          div.textContent = 'Message: ' + event.data;
          eventsDiv.appendChild(div);
        };
      </script>
    </body>
    </html>
  `);
});

// Start the server
app.listen(PORT, () => {
  const timestamp = new Date().toISOString();
  console.log(`[${timestamp}] Server running at http://localhost:${PORT}`);
  console.log(`[${timestamp}] SSE endpoint available at http://localhost:${PORT}/events`);
  console.log(`[${timestamp}] SSE test page available at http://localhost:${PORT}/test-sse`);
  console.log(`[${timestamp}] Waiting for client connections...`);
});
