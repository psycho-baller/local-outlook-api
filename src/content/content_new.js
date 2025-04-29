// Content script for interacting with Outlook Web Client
chrome.runtime.onMessage.addListener(function(request, sender, sendResponse) {
  if (request.action === 'sendEmail') {
    sendEmail(request.data)
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true; // Required for async sendResponse
  }
  
  if (request.action === 'getEventAttendees') {
    getEventAttendees(request.data)
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true; // Required for async sendResponse
  }
  
  if (request.action === 'getEventDetails') {
    getEventDetails(request.data)
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true; // Required for async sendResponse
  }
});

/**
 * Sends an email through Outlook Web Client
 * @param {Object} data - Email data
 * @param {string} data.to - Recipient email
 * @param {string} data.subject - Email subject
 * @param {string} data.body - Email body (HTML)
 * @returns {Promise<Object>} Result of the operation
 */
async function sendEmail(data) {
  try {
    // Check if we're on Outlook
    if (!window.location.href.includes('outlook.')) {
      throw new Error('Please navigate to Outlook Web Client');
    }

    // Click the New Message button
    const newMessageButton = await waitForElement('button[aria-label="New message"], button[aria-label="New mail"]');
    newMessageButton.click();
    
    // Wait for compose form to appear
    const toField = await waitForElement('div[aria-label="To"], input[aria-label="To"]');
    
    // Give the compose form time to fully initialize
    await sleep(1000);
    
    // Fill in recipient
    await fillRecipient(toField, data.to);
    
    // Fill in subject
    const subjectField = await waitForElement('input[aria-label="Subject"]');
    subjectField.focus();
    subjectField.value = data.subject;
    subjectField.dispatchEvent(new Event('input', { bubbles: true }));
    
    // Fill in body - find the editable div for the email body
    const bodyField = await waitForElement('div[role="textbox"][aria-label="Message body, press Alt+F10 to exit"], div[contenteditable="true"][aria-label="Message body, press Alt+F10 to exit"], div[role="textbox"][aria-label="Message body"], div[contenteditable="true"][aria-label="Message body"]');
    
    // Set HTML content
    bodyField.innerHTML = data.body;
    bodyField.dispatchEvent(new Event('input', { bubbles: true }));
    
    // Click Send button
    const sendButton = await waitForElement('button[aria-label="Send"], button[name="Send"]');
    sendButton.click();
    
    return { success: true };
  } catch (error) {
    console.error('Error sending email:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Fills in the recipient field
 * @param {Element} toField - The recipient input field
 * @param {string} email - Email address to enter
 */
async function fillRecipient(toField, email) {
  toField.focus();
  
  // Different approach based on whether it's a div or input
  if (toField.tagName.toLowerCase() === 'input') {
    toField.value = email;
    toField.dispatchEvent(new Event('input', { bubbles: true }));
  } else {
    // For contenteditable divs
    toField.textContent = email;
    toField.dispatchEvent(new Event('input', { bubbles: true }));
  }
  
  // Press Enter to confirm the recipient
  await sleep(500);
  toField.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  await sleep(500);
}

/**
 * Waits for an element to appear in the DOM
 * @param {string} selector - CSS selector for the element
 * @param {number} timeout - Timeout in milliseconds
 * @returns {Promise<Element>} The found element
 */
function waitForElement(selector, timeout = 10000) {
  return new Promise((resolve, reject) => {
    const startTime = Date.now();
    
    const checkElement = () => {
      const element = document.querySelector(selector);
      
      if (element) {
        resolve(element);
        return;
      }
      
      if (Date.now() - startTime > timeout) {
        reject(new Error(`Timeout waiting for element: ${selector}`));
        return;
      }
      
      setTimeout(checkElement, 100);
    };
    
    checkElement();
  });
}

/**
 * Sleep function for waiting
 * @param {number} ms - Milliseconds to sleep
 * @returns {Promise<void>}
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Gets detailed information about a specific event including attendees
 * @param {Object} data - Event data
 * @param {string} data.eventId - Event ID or unique identifier (can be part of the event title)
 * @returns {Promise<Object>} Result containing event details and attendees
 */
async function getEventDetails(data) {
  try {
    // Check if we're on Outlook
    if (!window.location.href.includes('outlook.')) {
      throw new Error('Please navigate to Outlook Web Client');
    }
    
    // Navigate to calendar view if not already there
    if (!window.location.href.includes('/calendar/')) {
      // Find and click the calendar navigation button
      const calendarButton = await waitForElement('button[aria-label="Calendar"], a[aria-label="Calendar"], [title="Calendar"]');
      calendarButton.click();
      
      // Wait for calendar to load
      await sleep(2000);
    }
    
    // Search for the event by title if eventId is provided
    if (data.eventId) {
      // Try to find the event in the current view
      let eventElements = document.querySelectorAll('[role="button"][aria-label*="' + data.eventId + '"], [role="gridcell"] [title*="' + data.eventId + '"], div[title*="' + data.eventId + '"]');
      
      // If not found, try to navigate to future months (up to 2 months ahead)
      if (eventElements.length === 0) {
        console.log('Event not found in current month, checking future months...');
        
        // Look for the next month button
        const nextMonthButton = document.querySelector('button[aria-label*="Next"], button[title*="Next"], button[aria-label*="next month"]');
        
        if (nextMonthButton) {
          // Try up to 2 future months
          for (let i = 0; i < 2; i++) {
            // Click next month
            nextMonthButton.click();
            
            // Wait for calendar to update
            await sleep(1000);
            
            // Look for the event again
            eventElements = document.querySelectorAll('[role="button"][aria-label*="' + data.eventId + '"], [role="gridcell"] [title*="' + data.eventId + '"], div[title*="' + data.eventId + '"]');
            
            if (eventElements.length > 0) {
              console.log(`Found event in month ${i + 1} ahead`);
              break;
            }
          }
          
          // If still not found, go back to the original month
          if (eventElements.length === 0) {
            const prevMonthButton = document.querySelector('button[aria-label*="Previous"], button[title*="Previous"], button[aria-label*="previous month"]');
            if (prevMonthButton) {
              for (let i = 0; i < 2; i++) {
                prevMonthButton.click();
                await sleep(500);
              }
            }
          }
        }
      }
      
      if (eventElements.length === 0) {
        throw new Error('Event not found: ' + data.eventId + ' (checked current and next 2 months)');
      }
      
      // Click on the first matching event
      eventElements[0].click();
      
      // Wait for event details to load
      await sleep(1000);
    } else {
      throw new Error('Event title or identifier is required');
    }
    
    // Extract event details
    const eventDetails = {};
    
    // Get event title
    const titleElement = document.querySelector('[role="heading"][aria-level="1"], [role="heading"][aria-level="2"], .ms-Dialog-title');
    if (titleElement) {
      eventDetails.title = titleElement.textContent.trim();
    }
    
    // Get organizer
    const organizerElement = document.querySelector('[aria-label*="Organizer"], [title*="Organizer"]');
    if (organizerElement) {
      eventDetails.organizer = organizerElement.textContent.trim();
    }
    
    // Get date and time
    const dateTimeElement = document.querySelector('[aria-label*="Date"], [title*="Date"], [aria-label*="Time"], [title*="Time"]');
    if (dateTimeElement) {
      eventDetails.datetime = dateTimeElement.textContent.trim();
    }
    
    // Get location
    const locationElement = document.querySelector('[aria-label*="Location"], [title*="Location"]');
    if (locationElement) {
      eventDetails.location = locationElement.textContent.trim();
    }
    
    // Get attendees (reuse the getEventAttendees logic)
    // Wait for the attendee list to appear
    const attendeeSection = await waitForElement('[aria-label="Attendees"], [role="region"][aria-label*="attendee"], [aria-label*="participant"]', 5000).catch(() => null);
    
    // Extract attendee information
    const attendees = [];
    
    if (attendeeSection) {
      // Look for accepted attendees
      const acceptedAttendeeElements = document.querySelectorAll('[aria-label*="Accepted"] [aria-label*="@"], [title*="Accepted"] [title*="@"]');
      
      for (const element of acceptedAttendeeElements) {
        // Extract name and email
        const fullText = element.getAttribute('aria-label') || element.getAttribute('title') || element.textContent;
        
        // Parse the text to extract name and email
        let name = '';
        let email = '';
        
        // Email is typically in format: name@domain.com
        const emailMatch = fullText.match(/[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/);
        if (emailMatch) {
          email = emailMatch[0];
          
          // Name is typically before the email
          const namePart = fullText.split(email)[0].trim();
          name = namePart.replace(/\s*\(.*?\)\s*$/, '').trim(); // Remove parentheses content if any
        } else {
          // If no email pattern found, use the text as name
          name = fullText.trim();
        }
        
        if (name || email) {
          attendees.push({ name, email });
        }
      }
    }
    
    // Add attendees to event details
    eventDetails.attendees = attendees;
    
    // Close the event details if needed
    const closeButton = document.querySelector('button[aria-label="Close"], button[aria-label="Back"], button.ms-Button--icon[aria-label*="close"]');
    if (closeButton) {
      closeButton.click();
    }
    
    return { 
      success: true,
      ...eventDetails,
      attendeeCount: attendees.length,
      eventId: data.eventId
    };
  } catch (error) {
    console.error('Error getting event details:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets the list of attendees who have accepted an invitation to a specific event
 * @param {Object} data - Event data
 * @param {string} data.eventId - Event ID or unique identifier (can be part of the event title)
 * @returns {Promise<Object>} Result containing the list of attendees
 */
async function getEventAttendees(data) {
  try {
    // Check if we're on Outlook
    if (!window.location.href.includes('outlook.')) {
      throw new Error('Please navigate to Outlook Web Client');
    }
    
    // Navigate to calendar view if not already there
    if (!window.location.href.includes('/calendar/')) {
      // Find and click the calendar navigation button
      const calendarButton = await waitForElement('button[aria-label="Calendar"], a[aria-label="Calendar"], [title="Calendar"]');
      calendarButton.click();
      
      // Wait for calendar to load
      await sleep(2000);
    }
    
    // Search for the event by title if eventId is provided
    if (data.eventId) {
      // Look for events with matching title or ID
      const eventElements = document.querySelectorAll('[role="button"][aria-label*="' + data.eventId + '"], [role="gridcell"] [title*="' + data.eventId + '"]');
      
      if (eventElements.length === 0) {
        throw new Error('Event not found: ' + data.eventId);
      }
      
      // Click on the first matching event
      eventElements[0].click();
      
      // Wait for event details to load
      await sleep(1000);
    } else {
      throw new Error('Event identifier is required');
    }
    
    // Wait for the attendee list to appear
    const attendeeSection = await waitForElement('[aria-label="Attendees"], [role="region"][aria-label*="attendee"], [aria-label*="participant"]');
    
    // Extract attendee information
    const attendees = [];
    
    // Look for accepted attendees
    const acceptedAttendeeElements = document.querySelectorAll('[aria-label*="Accepted"] [aria-label*="@"], [title*="Accepted"] [title*="@"]');
    
    for (const element of acceptedAttendeeElements) {
      // Extract name and email
      const fullText = element.getAttribute('aria-label') || element.getAttribute('title') || element.textContent;
      
      // Parse the text to extract name and email
      let name = '';
      let email = '';
      
      // Email is typically in format: name@domain.com
      const emailMatch = fullText.match(/[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/);
      if (emailMatch) {
        email = emailMatch[0];
        
        // Name is typically before the email
        const namePart = fullText.split(email)[0].trim();
        name = namePart.replace(/\s*\(.*?\)\s*$/, '').trim(); // Remove parentheses content if any
      } else {
        // If no email pattern found, use the text as name
        name = fullText.trim();
      }
      
      if (name || email) {
        attendees.push({ name, email });
      }
    }
    
    // Close the event details if needed
    const closeButton = document.querySelector('button[aria-label="Close"], button[aria-label="Back"], button.ms-Button--icon[aria-label*="close"]');
    if (closeButton) {
      closeButton.click();
    }
    
    return { 
      success: true, 
      attendees,
      count: attendees.length,
      eventId: data.eventId
    };
  } catch (error) {
    console.error('Error getting event attendees:', error);
    return { success: false, error: error.message };
  }
}
