// email-tab.js - Email composition tab functionality

// Initialize the email tab
function initEmailTab() {
  // Email tab elements
  const toInput = document.getElementById('to');
  const subjectInput = document.getElementById('subject');
  const bodyInput = document.getElementById('body');
  const bodyEditor = document.getElementById('body-editor');
  const sendBtn = document.getElementById('sendBtn');
  const clearBtn = document.getElementById('clearBtn');
  
  // Load any saved draft
  chrome.storage.local.get(['draftTo', 'draftSubject', 'draftBody'], function(result) {
    if (result.draftTo) toInput.value = result.draftTo;
    if (result.draftSubject) subjectInput.value = result.draftSubject;
    // Note: body content is loaded in common.js
  });
  
  // Save draft as user types
  toInput.addEventListener('input', saveDraft);
  subjectInput.addEventListener('input', saveDraft);
  
  function saveDraft() {
    // Get content from the editor
    const bodyContent = bodyInput.value || bodyEditor.innerHTML;
    
    chrome.storage.local.set({
      draftTo: toInput.value,
      draftSubject: subjectInput.value,
      draftBody: bodyContent
    });
  }
  
  // Send email
  sendBtn.addEventListener('click', function() {
    const to = toInput.value.trim();
    const subject = subjectInput.value.trim();
    const body = bodyInput.value.trim() || bodyEditor.innerHTML;

    // Validate inputs
    if (!to) {
      CommonUtils.showStatus('Please enter a recipient email address', 'error');
      return;
    }

    if (!CommonUtils.validateEmail(to)) {
      CommonUtils.showStatus('Please enter a valid email address', 'error');
      return;
    }

    if (!subject) {
      CommonUtils.showStatus('Please enter a subject', 'error');
      return;
    }

    if (!body) {
      CommonUtils.showStatus('Please enter a message body', 'error');
      return;
    }

    CommonUtils.showStatus('Preparing to send email...', 'pending');
    
    // First check if current tab is Outlook
    sendEmail(to, subject, body);
  });
  
  function sendEmail(to, subject, body) {
    chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
      const currentTab = tabs[0];
      const isOutlook = currentTab && currentTab.url &&
        (currentTab.url.includes('outlook.office.com') ||
          currentTab.url.includes('outlook.live.com'));

      if (isOutlook) {
        // Current tab is Outlook, send directly
        sendEmailToTab(currentTab.id, { to, subject, body });
      } else {
        // Current tab is not Outlook, try to find or create an Outlook tab
        chrome.runtime.sendMessage({ action: 'activateOutlookTab' }, function (result) {
          if (result && result.success) {
            // Give the tab time to load if it's new
            const delay = result.existing ? 500 : 3000;
            CommonUtils.showStatus(`${result.existing ? 'Switching to' : 'Opening'} Outlook tab...`, 'pending');

            setTimeout(() => {
              sendEmailToTab(result.tabId, { to, subject, body });
            }, delay);
          } else {
            CommonUtils.showStatus('Failed to open Outlook tab: ' + (result?.error || 'Unknown error'), 'error');
          }
        });
      }
    });
  }

  // Function to send email to a specific tab
  function sendEmailToTab(tabId, data) {
    CommonUtils.showStatus('Sending email...', 'pending');
    
    chrome.tabs.sendMessage(tabId, {
      action: 'sendEmail',
      data: data
    }, function(response) {
      if (chrome.runtime.lastError) {
        // Handle error when content script is not available
        CommonUtils.showStatus('Error: ' + chrome.runtime.lastError.message + '. Make sure Outlook is fully loaded.', 'error');
        return;
      }
      
      if (response && response.success) {
        CommonUtils.showStatus('Email sent successfully!', 'success');
        clearForm();
      } else {
        CommonUtils.showStatus(response?.error || 'Failed to send email. Make sure Outlook is open and loaded.', 'error');
      }
    });
  }

  // Clear form
  clearBtn.addEventListener('click', clearForm);

  function clearForm() {
    toInput.value = '';
    subjectInput.value = '';
    
    // Clear editor content
    bodyEditor.innerHTML = '';
    bodyInput.value = '';
    
    chrome.storage.local.remove(['draftTo', 'draftSubject', 'draftBody']);
    document.getElementById('status').textContent = '';
    document.getElementById('status').className = 'status';
  }
}

// Export functions
window.EmailTab = {
  init: initEmailTab
};
