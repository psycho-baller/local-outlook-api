// events-tab.js - Event details tab functionality

// Initialize the events tab
function initEventsTab() {
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
  
  // Load any saved event ID
  chrome.storage.local.get(['lastEventId'], function(result) {
    if (result.lastEventId) eventIdInput.value = result.lastEventId;
  });
  
  // Save event ID when user types
  eventIdInput.addEventListener('input', saveEventId);
  
  function saveEventId() {
    chrome.storage.local.set({
      lastEventId: eventIdInput.value
    });
  }
  
  // Get event details
  getEventBtn.addEventListener('click', function() {
    const eventId = eventIdInput.value.trim();
    
    if (!eventId) {
      CommonUtils.showStatus('Please enter an event title or ID', 'error');
      return;
    }
    
    CommonUtils.showStatus('Retrieving event details...', 'pending');
    eventResultDiv.style.display = 'none';
    
    // Find active tab
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      chrome.tabs.sendMessage(tabs[0].id, {
        action: 'getEventDetails',
        data: { eventId: eventId }
      }, function(response) {
        if (chrome.runtime.lastError) {
          CommonUtils.showStatus('Error: ' + chrome.runtime.lastError.message + '. Make sure Outlook is open.', 'error');
          return;
        }
        
        if (response && response.success) {
          displayEventDetails(response);
          CommonUtils.showStatus('Event details retrieved successfully', 'success');
        } else {
          CommonUtils.showStatus(response?.error || 'Failed to retrieve event details. Make sure Outlook is open.', 'error');
        }
      });
    });
  });
  
  // Display event details
  function displayEventDetails(data) {
    if (!data || !data.event) return;
    
    const event = data.event;
    
    // Set event details
    eventTitleSpan.textContent = event.title || 'N/A';
    eventOrganizerSpan.textContent = event.organizer || 'N/A';
    eventDatetimeSpan.textContent = event.datetime || 'N/A';
    eventLocationSpan.textContent = event.location || 'N/A';
    
    // Clear attendee list
    attendeeListDiv.innerHTML = '';
    
    // Add attendees
    if (event.attendees && event.attendees.length > 0) {
      attendeeCountSpan.textContent = `(${event.attendees.length})`;
      
      event.attendees.forEach(attendee => {
        const attendeeItem = document.createElement('div');
        attendeeItem.className = 'attendee-item';
        
        const nameSpan = document.createElement('div');
        nameSpan.textContent = attendee.name || 'Unknown';
        attendeeItem.appendChild(nameSpan);
        
        if (attendee.email) {
          const emailSpan = document.createElement('div');
          emailSpan.className = 'attendee-email';
          emailSpan.textContent = attendee.email;
          attendeeItem.appendChild(emailSpan);
        }
        
        attendeeListDiv.appendChild(attendeeItem);
      });
    } else {
      attendeeCountSpan.textContent = '(0)';
      
      const noAttendees = document.createElement('div');
      noAttendees.className = 'attendee-item';
      noAttendees.textContent = 'No attendees found';
      attendeeListDiv.appendChild(noAttendees);
    }
    
    // Show the result
    eventResultDiv.style.display = 'block';
  }
  
  // Copy all emails button
  copyEmailsBtn.addEventListener('click', function() {
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      chrome.tabs.sendMessage(tabs[0].id, {
        action: 'copyAttendeeEmails',
        data: { eventId: eventIdInput.value.trim() }
      }, function(response) {
        if (chrome.runtime.lastError) {
          CommonUtils.showStatus('Error: ' + chrome.runtime.lastError.message, 'error');
          return;
        }
        
        if (response && response.success) {
          // Copy to clipboard
          const emailsText = response.emails.join('; ');
          navigator.clipboard.writeText(emailsText).then(() => {
            CommonUtils.showStatus(`${response.emails.length} email addresses copied to clipboard`, 'success');
          }).catch(err => {
            CommonUtils.showStatus('Failed to copy to clipboard: ' + err, 'error');
          });
        } else {
          CommonUtils.showStatus(response?.error || 'Failed to get attendee emails', 'error');
        }
      });
    });
  });
}

// Export functions
window.EventsTab = {
  init: initEventsTab
};
