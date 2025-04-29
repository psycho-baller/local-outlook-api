/**
 * Calendar functionality for Outlook Web Client
 */

// waitForElement and sleep are loaded from utils.js

/**
 * Gets detailed information about a specific event including attendees
 * @param {Object} data - Event data
 * @param {string} data.eventId - Event ID or unique identifier (can be part of the event title)
 * @returns {Promise<Object>} Result containing event details and attendees
 */
async function getEventDetails(data) {
  try {
    // Check if we're in Outlook
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

    // Make sure we're in month view
    const monthButton = await waitForElement('button[name="Month"], [aria-label="Month view"], [title="Month"]');
    monthButton.click();
    await sleep(1000);

    // Search for the event by title if eventId is provided
    if (data.eventId) {
      // Try to find the event in the current view
      let eventElements = document.querySelectorAll(`[role="button"][aria-label*="${data.eventId}"], [role="gridcell"] [title*="${data.eventId}"], div[title*="${data.eventId}"]`);

      // If not found, try to navigate to future months (up to 2 months ahead)
      if (eventElements.length === 0) {
        console.log('Event not found in current month, checking future months...');

        // Look for the next month button
        const nextMonthButton = document.querySelector('button[aria-label^="Go to next month"], button[title*="Next"], button[aria-label*="next month"]');

        if (nextMonthButton) {
          // Try up to 2 future months
          for (let i = 0; i < 2; i++) {
            nextMonthButton.click();

            // Wait for calendar to update
            await sleep(1000);

            // Look for the event again
            eventElements = document.querySelectorAll(`[role="button"][aria-label*="${data.eventId}"], [role="gridcell"] [title*="${data.eventId}"], div[title*="${data.eventId}"]`);

            if (eventElements.length > 0) {
              console.log(`Found event in month ${i + 1} ahead`);
              break;
            }
          }

          // If still not found, go back to the original month
          if (eventElements.length === 0) {
            const prevMonthButton = document.querySelector('button[aria-label*="Previous"], button[title*="Previous"], button[aria-label*="previous month"]');
            for (let i = 0; i < 2; i++) {
              prevMonthButton.click();
              await sleep(500);
            }
          }
        }
      }

      if (eventElements.length === 0) {
        throw new Error('Event not found: ' + data.eventId + ' (checked current and next 2 months)');
      }

      // Double-click on the first matching event
      const element = eventElements[0];

      // Create and dispatch two click events in quick succession (double-click)
      const clickEvent = new MouseEvent('dblclick', {
        bubbles: true,
        cancelable: true,
        view: window
      });

      element.dispatchEvent(clickEvent);

      // Wait for event details to load
      await sleep(1500);

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

    // Find tracking info element
    const trackingElement = [...document.querySelectorAll('div')].find(d => d.textContent.trim() === 'Tracking');

    if (trackingElement) {
      console.log('Found Tracking element:', trackingElement);

      // try the following
      const moreOptionsElement = trackingElement.parentElement.querySelector('button[aria-label="More options"]');
      if (moreOptionsElement) {
        console.log('Found More options element:', moreOptionsElement);
        moreOptionsElement.click();

        await sleep(1500);
        const copyTrackingInfoElement = [...document.querySelectorAll('span')].find(s => s.textContent.trim() === 'Copy all tracking information');
        if (copyTrackingInfoElement) {
          copyTrackingInfoElement.click();
          await sleep(3000);
          navigator.clipboard.readText()
            .then(text => {
              console.log(text);
            })
            .catch(err => {
              console.error('Failed to read clipboard contents: ', err);
            });
        }
      }
    }
    
    // Get attendees (reuse the getEventAttendees logic)
    // Wait for the attendee list to appear
    const attendeeSection = await waitForElement('[aria-label="Attendees"], [role="region"][aria-label*="attendee"], [aria-label*="participant"]', 5000).catch(() => null);

    // Extract attendee information
    const attendees = [];

    if (attendeeSection) {
      // First, try to find the "Accepted" section by looking for a div with aria-label starting with "Accepted:"
      const acceptedHeaderElement = document.querySelector('div[aria-label^="Accepted:"], div[role="treeitem"][aria-label^="Accepted:"]');

      if (acceptedHeaderElement) {
        console.log('Found Accepted header element:', acceptedHeaderElement);

        // Get the parent element that contains all attendees
        const parentElement = acceptedHeaderElement.closest('div[role="tree"], div[role="group"]') || acceptedHeaderElement.parentElement;

        if (parentElement) {
          // Find all attendee elements that are siblings or children of the parent
          // Look for elements with aria-label containing "Open Profile Card" or elements with person names
          const attendeeElements = parentElement.querySelectorAll('div[role="treeitem"]:not([aria-label^="Accepted:"]), span[aria-label^="Open Profile Card"], div[class*="Persona"], div[aria-label*="Profile"]');

          console.log('Found attendee elements:', attendeeElements.length);

          for (const element of attendeeElements) {
            // Extract name and email
            let name = '';
            let email = '';

            // Try different attributes where the information might be stored
            const ariaLabel = element.getAttribute('aria-label') || '';
            const title = element.getAttribute('title') || '';
            const textContent = element.textContent || '';

            // Combine all possible sources of information
            const fullText = ariaLabel || title || textContent;

            // Email is typically in format: name@domain.com
            const emailMatch = fullText.match(/[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/);
            if (emailMatch) {
              email = emailMatch[0];

              // If we found an email, extract the name
              // Name might be in the format "Open Profile Card for Name, Status: Required"
              if (ariaLabel.includes('Open Profile Card for')) {
                name = ariaLabel.split('Open Profile Card for')[1].split(',')[0].trim();
              } else {
                // Name is typically before the email
                const namePart = fullText.split(email)[0].trim();
                name = namePart.replace(/\s*\(.*?\)\s*$/, '').trim(); // Remove parentheses content if any
              }
            } else {
              // If no email pattern found, look for name in the element or its children
              const nameElement = element.querySelector('[aria-label*="Profile"], [class*="Persona"]') || element;
              name = nameElement.textContent.trim();
            }

            if (name || email) {
              // Check if this attendee is already in the list to avoid duplicates
              const isDuplicate = attendees.some(a =>
                (a.email && a.email === email) ||
                (a.name && a.name === name)
              );

              if (!isDuplicate) {
                attendees.push({ name, email });
              }
            }
          }
        }
      }

      // If we didn't find any attendees with the new method, fall back to the original method
      if (attendees.length === 0) {
        console.log('Falling back to original attendee extraction method');

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
    // Check if we're in Outlook
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
      const eventElements = document.querySelectorAll(`[role="button"][aria-label*="${data.eventId}"], [role="gridcell"] [title*="${data.eventId}"], div[title*="${data.eventId}"]`);

      if (eventElements.length === 0) {
        throw new Error('Event not found: ' + data.eventId);
      }

      // Click on the first matching event
      eventElements[0].click();

      // Wait for event details to load
      await sleep(1000);
    } else {
      throw new Error('Event title or identifier is required');
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

// No export needed - these will be available in the global scope
