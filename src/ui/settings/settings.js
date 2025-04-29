document.addEventListener('DOMContentLoaded', function() {
  // DOM elements
  const sseEndpointInput = document.getElementById('sseEndpoint');
  const connectBtn = document.getElementById('connectBtn');
  const disconnectBtn = document.getElementById('disconnectBtn');
  const statusIndicator = document.getElementById('statusIndicator');
  const connectionStatus = document.getElementById('connectionStatus');
  const logContainer = document.getElementById('logContainer');
  const clearLogBtn = document.getElementById('clearLogBtn');
  const testEmailTo = document.getElementById('testEmailTo');
  const testEmailSubject = document.getElementById('testEmailSubject');
  const testEmailBody = document.getElementById('testEmailBody');
  const sendTestBtn = document.getElementById('sendTestBtn');
  
  // Connection state
  let isConnected = false;
  
  // Load saved endpoint URL
  chrome.storage.local.get(['sseEndpoint'], function(result) {
    if (result.sseEndpoint) {
      sseEndpointInput.value = result.sseEndpoint;
    }
  });
  
  // Check current connection status
  checkConnectionStatus();
  
  // Event listeners
  connectBtn.addEventListener('click', connectToSSE);
  disconnectBtn.addEventListener('click', disconnectFromSSE);
  clearLogBtn.addEventListener('click', clearLog);
  sendTestBtn.addEventListener('click', sendTestEmail);
  
  // Listen for SSE status updates from background script
  chrome.runtime.onMessage.addListener(function(message) {
    if (message.action === 'sseConnectionStatus') {
      updateConnectionStatus(message.status === 'connected');
      
      if (message.status === 'connected') {
        addLogEntry('Connected to SSE endpoint', 'success');
      } else if (message.status === 'failed') {
        addLogEntry('Failed to connect to SSE endpoint', 'error');
      }
    }
  });
  
  // Function to connect to SSE endpoint
  function connectToSSE() {
    const endpoint = sseEndpointInput.value.trim();
    
    if (!endpoint) {
      addLogEntry('Please enter a valid SSE endpoint URL', 'error');
      return;
    }
    
    // Save endpoint to storage
    chrome.storage.local.set({ sseEndpoint: endpoint });
    
    // Update UI to connecting state
    statusIndicator.className = 'status-indicator status-connecting';
    connectionStatus.textContent = 'Connecting...';
    connectBtn.disabled = true;
    disconnectBtn.disabled = true;
    
    addLogEntry(`Connecting to SSE endpoint: ${endpoint}`, 'info');
    
    // Send connection request to background script
    chrome.runtime.sendMessage(
      { action: 'connectToSSE', endpoint: endpoint },
      function(response) {
        if (response && response.success) {
          // Connection initiated successfully
          // The actual connection status will be updated via the message listener
          addLogEntry('Connection request sent successfully', 'info');
        } else {
          // Failed to initiate connection
          updateConnectionStatus(false);
          addLogEntry(`Failed to connect: ${response?.error || 'Unknown error'}`, 'error');
        }
      }
    );
  }
  
  // Function to disconnect from SSE endpoint
  function disconnectFromSSE() {
    chrome.runtime.sendMessage(
      { action: 'disconnectFromSSE' },
      function(response) {
        if (response && response.success) {
          updateConnectionStatus(false);
          addLogEntry('Disconnected from SSE endpoint', 'info');
        } else {
          addLogEntry(`Failed to disconnect: ${response?.error || 'Unknown error'}`, 'error');
        }
      }
    );
  }
  
  // Function to check current connection status
  function checkConnectionStatus() {
    chrome.runtime.sendMessage(
      { action: 'checkSSEStatus' },
      function(response) {
        if (response) {
          updateConnectionStatus(response.isConnected);
          
          if (response.isConnected && response.endpoint) {
            addLogEntry(`Currently connected to: ${response.endpoint}`, 'info');
            sseEndpointInput.value = response.endpoint;
          }
        }
      }
    );
  }
  
  // Function to update connection status UI
  function updateConnectionStatus(connected) {
    isConnected = connected;
    
    if (connected) {
      statusIndicator.className = 'status-indicator status-connected';
      connectionStatus.textContent = 'Connected';
      connectBtn.disabled = true;
      disconnectBtn.disabled = false;
      sendTestBtn.disabled = false;
    } else {
      statusIndicator.className = 'status-indicator status-disconnected';
      connectionStatus.textContent = 'Disconnected';
      connectBtn.disabled = false;
      disconnectBtn.disabled = true;
      sendTestBtn.disabled = true;
    }
  }
  
  // Function to send a test email
  function sendTestEmail() {
    const to = testEmailTo.value.trim();
    const subject = testEmailSubject.value.trim() || 'Test Email';
    const body = testEmailBody.value.trim() || 'This is a test email sent from the Outlook Email Assistant.';
    
    if (!to) {
      addLogEntry('Please enter a recipient email address', 'error');
      return;
    }
    
    addLogEntry(`Sending test email to: ${to}`, 'info');
    
    // Create a test email instruction
    const emailInstruction = {
      to: to,
      subject: subject,
      body: body
    };
    
    // Send directly to the processEmailInstruction function in the background script
    chrome.runtime.sendMessage(
      { 
        action: 'testEmailInstruction', 
        instruction: emailInstruction 
      },
      function(response) {
        if (response && response.success) {
          addLogEntry('Test email sent successfully', 'success');
        } else {
          addLogEntry(`Failed to send test email: ${response?.error || 'Unknown error'}`, 'error');
        }
      }
    );
  }
  
  // Function to add log entry
  function addLogEntry(message, type = 'info') {
    const entry = document.createElement('div');
    entry.className = `log-entry log-${type}`;
    
    const timestamp = new Date().toLocaleTimeString();
    entry.textContent = `[${timestamp}] ${message}`;
    
    logContainer.appendChild(entry);
    logContainer.scrollTop = logContainer.scrollHeight;
  }
  
  // Function to clear log
  function clearLog() {
    logContainer.innerHTML = '';
  }
});
