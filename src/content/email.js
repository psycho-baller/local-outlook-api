/**
 * Email functionality for Outlook Web Client
 */

// waitForElement and sleep are loaded from utils.js

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
    // Check if we're in Outlook
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
    
    // Try to save as draft if sending fails
    try {
      const closeButton = document.querySelector('button[aria-label="Discard"], button[aria-label="Close"]');
      if (closeButton) {
        closeButton.click();
        await sleep(500);
        
        // Click "Save" if prompted
        const saveButton = document.querySelector('button[aria-label="Save"], button[name="Save"]');
        if (saveButton) {
          saveButton.click();
          return { success: false, error: error.message, draftSaved: true };
        }
      }
    } catch (draftError) {
      console.error('Error saving draft:', draftError);
    }
    
    return { success: false, error: error.message };
  }
}

/**
 * Fills in the recipient field
 * @param {Element} toField - The recipient input field
 * @param {string} email - Email address to enter
 * @returns {Promise<void>}
 */
async function fillRecipient(toField, email) {
  toField.focus();

  // Different approach based on whether it's a div or input
  if (toField.tagName.toLowerCase() === 'input') {
    toField.value = email;
    toField.dispatchEvent(new Event('input', { bubbles: true }));
  } else {
    toField.textContent = email;
    toField.dispatchEvent(new Event('input', { bubbles: true }));
  }

  // Press Enter to confirm the recipient
  await sleep(500);
  toField.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  await sleep(500);
}

// No export needed - this will be available in the global scope
