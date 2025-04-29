/**
 * Content script for interacting with Outlook Web Client
 * This file serves as the entry point and message router for the extension
 */

// Functions from other scripts are already available in the global scope
// Browser Extension / manifest v3 does not support modules, hence this workaround
// - utils.js provides: waitForElement, sleep
// - email.js provides: sendEmail
// - calendar.js provides: getEventDetails, getEventAttendees

// Listen for messages from the extension's background script
chrome.runtime.onMessage.addListener(function (request, sender, sendResponse) {
  // Handle email-related messages
  if (request.action === 'sendEmail') {
    sendEmail(request.data)
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true; // Required for async sendResponse
  }

  // Handle event-related messages
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

  // If no handler matched, return false
  return false;
});
