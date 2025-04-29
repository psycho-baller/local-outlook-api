document.addEventListener('DOMContentLoaded', function() {
  // Email tab elements
  const toInput = document.getElementById('to');
  const subjectInput = document.getElementById('subject');
  const bodyInput = document.getElementById('body');
  const sendBtn = document.getElementById('sendBtn');
  const clearBtn = document.getElementById('clearBtn');
  
  // Event tab elements
  const eventIdInput = document.getElementById('eventId');
  const getEventBtn = document.getElementById('getEventBtn');
  const eventResultDiv = document.getElementById('event-result');
  const eventTitleSpan = document.getElementById('event-title');
  const eventOrganizerSpan = document.getElementById('event-organizer');
  const eventDatetimeSpan = document.getElementById('event-datetime');
  const eventLocationSpan = document.getElementById('event-location');
  const attendeeListDiv = document.getElementById('attendee-list');
  const attendeeCountSpan = document.getElementById('attendee-count');
  const copyEmailsBtn = document.getElementById('copyEmailsBtn');
  
  // Tab navigation elements
  const tabs = document.querySelectorAll('.tab');
  const tabContents = document.querySelectorAll('.tab-content');
  
  // Status element
  const statusDiv = document.getElementById('status');

  // Load any saved draft and event ID
  chrome.storage.local.get(['draftTo', 'draftSubject', 'draftBody', 'lastEventId'], function(result) {
    if (result.draftTo) toInput.value = result.draftTo;
    if (result.draftSubject) subjectInput.value = result.draftSubject;
    if (result.draftBody) bodyInput.value = result.draftBody;
    if (result.lastEventId) eventIdInput.value = result.lastEventId;
  });
  
  // Tab switching functionality
  tabs.forEach(tab => {
    tab.addEventListener('click', function() {
      // Remove active class from all tabs and contents
      tabs.forEach(t => t.classList.remove('active'));
      tabContents.forEach(c => c.classList.remove('active'));
      
      // Add active class to clicked tab and corresponding content
      const tabId = this.getAttribute('data-tab');
      this.classList.add('active');
      document.getElementById(tabId + '-tab').classList.add('active');
      
      // Clear status when switching tabs
      showStatus('');
    });
  });

  // Save draft as user types
  function saveDraft() {
    chrome.storage.local.set({
      draftTo: toInput.value,
      draftSubject: subjectInput.value,
      draftBody: bodyInput.value
    });
  }
  
  // Save event ID when user types
  function saveEventId() {
    chrome.storage.local.set({
      lastEventId: eventIdInput.value
    });
  }

  toInput.addEventListener('input', saveDraft);
  subjectInput.addEventListener('input', saveDraft);
  bodyInput.addEventListener('input', saveDraft);
  eventIdInput.addEventListener('input', saveEventId);

  // Send email
  sendBtn.addEventListener('click', function() {
    const to = toInput.value.trim();
    const subject = subjectInput.value.trim();
    const body = bodyInput.value.trim();

    // Validate inputs
    if (!to) {
      showStatus('Please enter a recipient email address', 'error');
      return;
    }

    if (!validateEmail(to)) {
      showStatus('Please enter a valid email address', 'error');
      return;
    }

    if (!subject) {
      showStatus('Please enter a subject', 'error');
      return;
    }

    if (!body) {
      showStatus('Please enter a message body', 'error');
      return;
    }

    showStatus('Preparing to send email...', 'pending');
    
    // First check if current tab is Outlook
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      const currentTab = tabs[0];
      const isOutlook = currentTab && currentTab.url && 
                       (currentTab.url.includes('outlook.office.com') || 
                        currentTab.url.includes('outlook.live.com'));
      
      if (isOutlook) {
        // Current tab is Outlook, send directly
        sendEmailToTab(currentTab.id, { to, subject, body });
      } else {
        // Current tab is not Outlook, try to find or create an Outlook tab
        chrome.runtime.sendMessage({ action: 'activateOutlookTab' }, function(result) {
          if (result && result.success) {
            // Give the tab time to load if it's new
            const delay = result.existing ? 500 : 3000;
            showStatus(`${result.existing ? 'Switching to' : 'Opening'} Outlook tab...`, 'pending');
            
            setTimeout(() => {
              sendEmailToTab(result.tabId, { to, subject, body });
            }, delay);
          } else {
            showStatus('Failed to open Outlook tab: ' + (result?.error || 'Unknown error'), 'error');
          }
        });
      }
    });
  });
  
  // Function to send email to a specific tab
  function sendEmailToTab(tabId, data) {
    showStatus('Sending email...', 'pending');
    
    chrome.tabs.sendMessage(tabId, {
      action: 'sendEmail',
      data: data
    }, function(response) {
      if (chrome.runtime.lastError) {
        // Handle error when content script is not available
        showStatus('Error: ' + chrome.runtime.lastError.message + '. Make sure Outlook is fully loaded.', 'error');
        return;
      }
      
      if (response && response.success) {
        showStatus('Email sent successfully!', 'success');
        clearForm();
      } else {
        showStatus(response?.error || 'Failed to send email. Make sure Outlook is open and loaded.', 'error');
      }
    });
  }

  // Clear form
  clearBtn.addEventListener('click', clearForm);

  function clearForm() {
    toInput.value = '';
    subjectInput.value = '';
    bodyInput.value = '';
    chrome.storage.local.remove(['draftTo', 'draftSubject', 'draftBody']);
    statusDiv.textContent = '';
    statusDiv.className = 'status';
  }
  
  // Get event details
  getEventBtn.addEventListener('click', function() {
    const eventId = eventIdInput.value.trim();
    
    if (!eventId) {
      showStatus('Please enter an event title or ID', 'error');
      return;
    }
    
    showStatus('Retrieving event details...', 'pending');
    eventResultDiv.style.display = 'none';
    
    // Find active tab
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      chrome.tabs.sendMessage(tabs[0].id, {
        action: 'getEventDetails',
        data: { eventId: eventId }
      }, function(response) {
        if (chrome.runtime.lastError) {
          showStatus('Error: ' + chrome.runtime.lastError.message + '. Make sure Outlook is open.', 'error');
          return;
        }
        
        if (response && response.success) {
          displayEventDetails(response);
          showStatus('Event details retrieved successfully', 'success');
        } else {
          showStatus(response?.error || 'Failed to retrieve event details. Make sure Outlook is open.', 'error');
        }
      });
    });
  });
  
  // Display event details
  function displayEventDetails(data) {
    // Set event details
    eventTitleSpan.textContent = data.title || 'N/A';
    eventOrganizerSpan.textContent = data.organizer || 'N/A';
    eventDatetimeSpan.textContent = data.datetime || 'N/A';
    eventLocationSpan.textContent = data.location || 'N/A';
    
    // Set attendee count
    const attendeeCount = data.attendees ? data.attendees.length : 0;
    attendeeCountSpan.textContent = `(${attendeeCount})`;
    
    // Clear previous attendees
    attendeeListDiv.innerHTML = '';
    
    // Add attendees to the list
    if (data.attendees && data.attendees.length > 0) {
      data.attendees.forEach(attendee => {
        const attendeeItem = document.createElement('div');
        attendeeItem.className = 'attendee-item';
        
        const nameSpan = document.createElement('div');
        nameSpan.textContent = attendee.name || 'No Name';
        
        const emailSpan = document.createElement('div');
        emailSpan.className = 'attendee-email';
        emailSpan.textContent = attendee.email || 'No Email';
        
        attendeeItem.appendChild(nameSpan);
        attendeeItem.appendChild(emailSpan);
        attendeeListDiv.appendChild(attendeeItem);
      });
    } else {
      const noAttendeesItem = document.createElement('div');
      noAttendeesItem.className = 'attendee-item';
      noAttendeesItem.textContent = 'No accepted attendees found';
      attendeeListDiv.appendChild(noAttendeesItem);
    }
    
    // Show the result div
    eventResultDiv.style.display = 'block';
  }
  
  // Copy all emails button
  copyEmailsBtn.addEventListener('click', function() {
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      chrome.tabs.sendMessage(tabs[0].id, {
        action: 'getEventDetails',
        data: { eventId: eventIdInput.value.trim() }
      }, function(response) {
        if (response && response.success && response.attendees && response.attendees.length > 0) {
          const emails = response.attendees
            .map(a => a.email)
            .filter(e => e)
            .join(', ');
          
          if (emails) {
            navigator.clipboard.writeText(emails).then(function() {
              showStatus('All emails copied to clipboard!', 'success');
            }, function() {
              showStatus('Failed to copy emails to clipboard', 'error');
            });
          } else {
            showStatus('No valid emails to copy', 'error');
          }
        } else {
          showStatus('No attendee emails available', 'error');
        }
      });
    });
  });

  function showStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
  }

  function validateEmail(email) {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
  }
});
