// Background script for the Outlook Email Assistant extension

// Configuration for SSE connection
let sseConfig = {
  endpoint: null,      // SSE endpoint URL
  reconnectDelay: 3000, // Delay between reconnection attempts in ms
  isConnected: false,   // Connection status
  eventSource: null,    // EventSource instance
  connectionAttempts: 0, // Number of connection attempts
  maxConnectionAttempts: 5 // Maximum number of connection attempts before giving up
};

// Open the side panel when the extension icon is clicked
chrome.action.onClicked.addListener((tab) => {
  // Check if the current URL is Outlook
  if (tab.url.includes('outlook.office.com') || tab.url.includes('outlook.live.com')) {
    // Open the side panel
    chrome.sidePanel.open({ tabId: tab.id });
  } else {
    // Notify the user to navigate to Outlook
    chrome.action.setPopup({ popup: 'popup.html' });
  }
});

// Function to find an Outlook tab in all windows
async function findOutlookTab() {
  return new Promise((resolve) => {
    chrome.tabs.query({}, (tabs) => {
      const outlookTab = tabs.find(tab => 
        tab.url && (tab.url.includes('outlook.office.com') || tab.url.includes('outlook.live.com'))
      );
      resolve(outlookTab || null);
    });
  });
}

// Function to activate an Outlook tab or create one if none exists
async function activateOrCreateOutlookTab() {
  const outlookTab = await findOutlookTab();
  
  if (outlookTab) {
    // Activate the existing Outlook tab
    await chrome.tabs.update(outlookTab.id, { active: true });
    await chrome.windows.update(outlookTab.windowId, { focused: true });
    return { success: true, tabId: outlookTab.id, existing: true };
  } else {
    // Create a new Outlook tab
    try {
      const newTab = await chrome.tabs.create({ url: 'https://outlook.office.com/mail/' });
      return { success: true, tabId: newTab.id, existing: false };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }
}

// Function to connect to SSE endpoint
async function connectToSSE(endpoint) {
  // Store the endpoint URL
  sseConfig.endpoint = endpoint;
  
  // Close any existing connection
  if (sseConfig.eventSource) {
    sseConfig.eventSource.close();
    sseConfig.eventSource = null;
  }
  
  try {
    // Create a new EventSource connection
    const eventSource = new EventSource(endpoint);
    sseConfig.eventSource = eventSource;
    sseConfig.connectionAttempts = 0;
    
    // Set up event listeners
    eventSource.onopen = (event) => {
      console.log('SSE connection established:', endpoint);
      sseConfig.isConnected = true;
      // Notify any interested components that we're connected
      chrome.runtime.sendMessage({ action: 'sseConnectionStatus', status: 'connected' });
    };
    
    // Handle incoming email instructions
    eventSource.addEventListener('email-instruction', async (event) => {
      try {
        const data = JSON.parse(event.data);
        console.log('Received email instruction:', data);
        
        // Process the email instruction
        const result = await processEmailInstruction(data);
        
        // Send the result back to the server if callback URL is provided
        if (data.callbackUrl) {
          console.log('Sending email result to callback URL:', data.callbackUrl);
          fetch(data.callbackUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              requestId: data.requestId || 'unknown',
              ...result
            })
          }).catch(error => {
            console.error('Error sending email result back to server:', error);
          });
        }
      } catch (error) {
        console.error('Error processing email instruction:', error);
        
        // Send error to callback URL if provided
        if (data && data.callbackUrl) {
          fetch(data.callbackUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              requestId: data.requestId || 'unknown',
              success: false,
              error: error.message || 'Unknown error'
            })
          }).catch(callbackError => {
            console.error('Error sending error to callback URL:', callbackError);
          });
        }
      }
    });
    
    // Handle event attendee requests
    eventSource.addEventListener('get-event-attendees', async (event) => {
      try {
        const data = JSON.parse(event.data);
        console.log('Received event attendees request:', data);
        
        // Process the event attendees request
        const result = await processEventAttendeesRequest(data);
        
        // Send the result back to the server
        const callbackUrl = data.callbackUrl || (sseConfig.endpoint.replace('/events', '/event-attendees-result'));
        console.log('Sending attendees result to callback URL:', callbackUrl);
        
        fetch(callbackUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            requestId: data.requestId || 'unknown',
            ...result
          })
        }).catch(error => {
          console.error('Error sending attendees result back to server:', error);
        });
      } catch (error) {
        console.error('Error processing event attendees request:', error);
        
        // Send error to callback URL if provided
        if (data && (data.callbackUrl || sseConfig.endpoint)) {
          const callbackUrl = data.callbackUrl || (sseConfig.endpoint.replace('/events', '/event-attendees-result'));
          
          fetch(callbackUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              requestId: data.requestId || 'unknown',
              success: false,
              error: error.message || 'Unknown error'
            })
          }).catch(callbackError => {
            console.error('Error sending error to callback URL:', callbackError);
          });
        }
      }
    });
    
    // Handle general messages
    eventSource.onmessage = (event) => {
      console.log('Received message from SSE server:', event.data);
    };
    
    // Handle errors
    eventSource.onerror = (error) => {
      console.error('SSE connection error:', error);
      sseConfig.isConnected = false;
      sseConfig.eventSource.close();
      sseConfig.eventSource = null;
      
      // Attempt to reconnect
      sseConfig.connectionAttempts++;
      if (sseConfig.connectionAttempts < sseConfig.maxConnectionAttempts) {
        console.log(`Attempting to reconnect (${sseConfig.connectionAttempts}/${sseConfig.maxConnectionAttempts})...`);
        setTimeout(() => {
          connectToSSE(sseConfig.endpoint);
        }, sseConfig.reconnectDelay);
      } else {
        console.error('Max reconnection attempts reached. Giving up.');
        chrome.runtime.sendMessage({ action: 'sseConnectionStatus', status: 'failed' });
      }
    };
    
    return true;
  } catch (error) {
    console.error('Failed to connect to SSE endpoint:', error);
    return false;
  }
}

// Process email instructions received from SSE
async function processEmailInstruction(instruction) {
  // Validate the instruction
  if (!instruction.to || !instruction.subject || !instruction.body) {
    console.error('Invalid email instruction:', instruction);
    return { success: false, error: 'Invalid email instruction', requestId: instruction.requestId };
  }
  
  try {
    // Find or create an Outlook tab
    const outlookTabResult = await activateOrCreateOutlookTab();
    
    if (!outlookTabResult.success) {
      throw new Error('Failed to open Outlook tab: ' + outlookTabResult.error);
    }
    
    // Give the tab time to load if it's new
    const delay = outlookTabResult.existing ? 500 : 3000;
    await new Promise(resolve => setTimeout(resolve, delay));
    
    // Send the email instruction to the content script
    return new Promise((resolve) => {
      chrome.tabs.sendMessage(outlookTabResult.tabId, {
        action: 'sendEmail',
        data: {
          to: instruction.to,
          subject: instruction.subject,
          body: instruction.body
        }
      }, (response) => {
        if (chrome.runtime.lastError) {
          console.error('Error sending message to content script:', chrome.runtime.lastError);
          resolve({ 
            success: false, 
            error: chrome.runtime.lastError.message,
            requestId: instruction.requestId
          });
          return;
        }
        
        if (response && response.success) {
          console.log('Email sent successfully via SSE instruction');
          resolve({ 
            success: true,
            requestId: instruction.requestId
          });
        } else {
          console.error('Failed to send email:', response?.error || 'Unknown error');
          resolve({ 
            success: false, 
            error: response?.error || 'Unknown error',
            requestId: instruction.requestId
          });
        }
      });
    });
  } catch (error) {
    console.error('Error processing email instruction:', error);
    return { 
      success: false, 
      error: error.message,
      requestId: instruction.requestId
    };
  }
}

// Process event attendees requests received from SSE
async function processEventAttendeesRequest(request) {
  // Validate the request
  if (!request.eventId) {
    console.error('Invalid event attendees request:', request);
    return { success: false, error: 'Event ID is required', requestId: request.requestId };
  }
  
  try {
    // Find or create an Outlook tab
    const outlookTabResult = await activateOrCreateOutlookTab();
    
    if (!outlookTabResult.success) {
      throw new Error('Failed to open Outlook tab: ' + outlookTabResult.error);
    }
    
    // Give the tab time to load if it's new
    const delay = outlookTabResult.existing ? 500 : 3000;
    await new Promise(resolve => setTimeout(resolve, delay));
    
    // Send the email instruction to the content script
    return new Promise((resolve) => {
      chrome.tabs.sendMessage(outlookTabResult.tabId, {
        action: 'sendEmail',
        data: {
          to: instruction.to,
          subject: instruction.subject,
          body: instruction.body
        }
      }, (response) => {
        if (chrome.runtime.lastError) {
          console.error('Error sending message to content script:', chrome.runtime.lastError);
          resolve({ success: false, error: chrome.runtime.lastError.message });
          return;
        }
        
        if (response && response.success) {
          console.log('Email sent successfully via SSE instruction');
          resolve({ success: true });
        } else {
          console.error('Failed to send email:', response?.error || 'Unknown error');
          resolve({ success: false, error: response?.error || 'Unknown error' });
        }
      });
    });
  } catch (error) {
    console.error('Error processing email instruction:', error);
    return { success: false, error: error.message };
  }
}

async function processEventAttendeesRequest(request) {
  try {
    // Find or create an Outlook tab
    const outlookTabResult = await activateOrCreateOutlookTab();
    
    if (!outlookTabResult.success) {
      throw new Error('Failed to open Outlook tab: ' + outlookTabResult.error);
    }
    
    // Give the tab time to load if it's new
    const delay = outlookTabResult.existing ? 500 : 3000;
    await new Promise(resolve => setTimeout(resolve, delay));
    
    // Send the event attendees request to the content script
    return new Promise((resolve) => {
      chrome.tabs.sendMessage(outlookTabResult.tabId, {
        action: 'getEventAttendees',
        data: {
          eventId: request.eventId
        }
      }, (response) => {
        if (chrome.runtime.lastError) {
          console.error('Error sending message to content script:', chrome.runtime.lastError);
          resolve({ success: false, error: chrome.runtime.lastError.message });
          return;
        }
        
        if (response && response.success) {
          console.log('Event attendees retrieved successfully:', response.attendees);
          resolve({
            success: true,
            attendees: response.attendees,
            count: response.count,
            eventId: request.eventId,
            requestId: request.requestId
          });
        } else {
          console.error('Failed to get event attendees:', response?.error || 'Unknown error');
          resolve({ success: false, error: response?.error || 'Unknown error' });
        }
      });
    });
  } catch (error) {
    console.error('Error processing email instruction:', error);
    return { success: false, error: error.message };
  }
}

// Listen for messages from the side panel or popup
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'checkOutlook') {
    // Check if the active tab is Outlook
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      const isOutlook = tabs[0]?.url?.includes('outlook.');
      sendResponse({ isOutlook });
    });
    return true; // Required for async sendResponse
  }
  
  if (message.action === 'findOutlookTab') {
    findOutlookTab().then(tab => {
      sendResponse({ tab });
    });
    return true; // Required for async sendResponse
  }
  
  if (message.action === 'activateOutlookTab') {
    activateOrCreateOutlookTab().then(result => {
      sendResponse(result);
    });
    return true; // Required for async sendResponse
  }
  
  // Handle SSE connection requests
  if (message.action === 'connectToSSE') {
    if (!message.endpoint) {
      sendResponse({ success: false, error: 'No SSE endpoint provided' });
      return true;
    }
    
    connectToSSE(message.endpoint).then(success => {
      sendResponse({ success, isConnected: sseConfig.isConnected });
    }).catch(error => {
      sendResponse({ success: false, error: error.message });
    });
    return true; // Required for async sendResponse
  }
  
  // Handle SSE disconnection requests
  if (message.action === 'disconnectFromSSE') {
    if (sseConfig.eventSource) {
      sseConfig.eventSource.close();
      sseConfig.eventSource = null;
      sseConfig.isConnected = false;
      sendResponse({ success: true });
    } else {
      sendResponse({ success: false, error: 'No active SSE connection' });
    }
    return true;
  }
  
  // Check SSE connection status
  if (message.action === 'checkSSEStatus') {
    sendResponse({
      isConnected: sseConfig.isConnected,
      endpoint: sseConfig.endpoint
    });
    return true;
  }
  
  // Handle test email instructions from the settings page
  if (message.action === 'testEmailInstruction' && message.instruction) {
    processEmailInstruction(message.instruction).then(result => {
      sendResponse(result);
    }).catch(error => {
      sendResponse({ success: false, error: error.message });
    });
    return true; // Required for async sendResponse
  }
});
