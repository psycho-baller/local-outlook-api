// common.js - Shared utilities and functions

// Email validation function
function validateEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

// Show status message
function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = 'status ' + (type || '');
}

// Rich text editor functionality
function initRichTextEditors() {
  // Get HTML toggle buttons from toolbar
  const bodyHtmlToggleBtn = document.getElementById('body-html-toggle-btn');
  const bulkBodyHtmlToggleBtn = document.getElementById('bulk-body-html-toggle-btn');
  
  // Set up the regular email editor
  if (document.getElementById('body-editor')) {
    setupEditor(
      document.getElementById('body-editor'),
      document.getElementById('body-toolbar'),
      document.getElementById('body-format-select'),
      document.getElementById('body'),
      document.getElementById('body-html-toggle'),
      bodyHtmlToggleBtn,
      function() {
        // Save draft callback
        const toInput = document.getElementById('to');
        const subjectInput = document.getElementById('subject');
        const bodyInput = document.getElementById('body');
        const bodyEditor = document.getElementById('body-editor');
        
        const bodyContent = bodyInput.value || bodyEditor.innerHTML;
        
        chrome.storage.local.set({
          draftTo: toInput.value,
          draftSubject: subjectInput.value,
          draftBody: bodyContent
        });
      }
    );
  }
  
  // Set up the bulk email editor
  if (document.getElementById('bulk-body-editor')) {
    setupEditor(
      document.getElementById('bulk-body-editor'),
      document.getElementById('bulk-body-toolbar'),
      document.getElementById('bulk-body-format-select'),
      document.getElementById('bulk-body'),
      document.getElementById('bulk-body-html-toggle'),
      bulkBodyHtmlToggleBtn,
      function() {
        // Save bulk data callback
        const bulkSubjectInput = document.getElementById('bulk-subject');
        const bulkBodyInput = document.getElementById('bulk-body');
        const bulkBodyEditor = document.getElementById('bulk-body-editor');
        const bulkRecipientsInput = document.getElementById('bulk-recipients');
        const formatRadios = document.querySelectorAll('input[name="data-format"]');
        
        let selectedFormat = 'auto';
        formatRadios.forEach(radio => {
          if (radio.checked) {
            selectedFormat = radio.value;
          }
        });
        
        const bodyContent = bulkBodyInput.value || bulkBodyEditor.innerHTML;
        
        chrome.storage.local.set({
          bulkSubject: bulkSubjectInput.value,
          bulkBody: bodyContent,
          bulkRecipients: bulkRecipientsInput.value,
          bulkFormat: selectedFormat
        });
      }
    );
  }
  
  // Load saved content
  chrome.storage.local.get(['draftBody', 'bulkBody'], function(result) {
    if (result.draftBody && document.getElementById('body-editor')) {
      document.getElementById('body-editor').innerHTML = result.draftBody;
      document.getElementById('body').value = result.draftBody;
    }
    
    if (result.bulkBody && document.getElementById('bulk-body-editor')) {
      document.getElementById('bulk-body-editor').innerHTML = result.bulkBody;
      document.getElementById('bulk-body').value = result.bulkBody;
    }
  });
}

// Set up a rich text editor
function setupEditor(editorElement, toolbarElement, formatSelect, hiddenInput, htmlToggle, htmlToggleBtn, saveCallback) {
  // Make sure the editor is focused when clicked
  editorElement.addEventListener('focus', function() {
    // Set focus to the end of the content
    const range = document.createRange();
    const selection = window.getSelection();
    range.selectNodeContents(editorElement);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  });
  
  // Save content when editor changes
  editorElement.addEventListener('input', function() {
    hiddenInput.value = editorElement.innerHTML;
    saveCallback();
  });
  
  // Set up toolbar buttons
  const buttons = toolbarElement.querySelectorAll('button[data-command]');
  buttons.forEach(button => {
    button.addEventListener('click', function() {
      const command = this.dataset.command;
      const value = this.dataset.value || null;
      
      // Special handling for createLink
      if (command === 'createLink') {
        const url = prompt('Enter the URL:', 'http://');
        if (url) {
          document.execCommand(command, false, url);
        }
      } else {
        document.execCommand(command, false, value);
      }
      
      // Focus back to the editor
      editorElement.focus();
      
      // Save changes
      hiddenInput.value = editorElement.innerHTML;
      saveCallback();
    });
  });
  
  // Set up format select
  formatSelect.addEventListener('change', function() {
    const format = this.value;
    document.execCommand('formatBlock', false, format);
    
    // Focus back to the editor
    editorElement.focus();
    
    // Save changes
    hiddenInput.value = editorElement.innerHTML;
    saveCallback();
  });
  
  // Handle HTML toggle checkbox (hidden but still functional)
  htmlToggle.addEventListener('change', function() {
    toggleHtmlMode(editorElement, hiddenInput, this.checked, htmlToggleBtn, saveCallback);
  });
  
  // Handle HTML toggle button in toolbar
  htmlToggleBtn.addEventListener('click', function() {
    const isHtmlMode = htmlToggleBtn.classList.contains('active');
    htmlToggle.checked = !isHtmlMode;
    toggleHtmlMode(editorElement, hiddenInput, !isHtmlMode, htmlToggleBtn, saveCallback);
  });
}

// Toggle between rich editor and HTML mode
function toggleHtmlMode(editorElement, hiddenInput, showHtml, htmlToggleBtn, saveCallback) {
  if (showHtml) {
    // Switch to HTML mode
    const content = editorElement.innerHTML;
    
    // Create a textarea for HTML editing if it doesn't exist
    let htmlTextarea = document.getElementById(editorElement.id + '-html');
    if (!htmlTextarea) {
      htmlTextarea = document.createElement('textarea');
      htmlTextarea.id = editorElement.id + '-html';
      htmlTextarea.className = 'html-mode';
      
      // Update content when HTML is edited
      htmlTextarea.addEventListener('input', function() {
        editorElement.innerHTML = this.value;
        hiddenInput.value = this.value;
        saveCallback();
      });
      
      editorElement.insertAdjacentElement('afterend', htmlTextarea);
    }
    
    htmlTextarea.value = content;
    editorElement.style.display = 'none';
    htmlTextarea.style.display = 'block';
    
    // Update the toggle button appearance
    htmlToggleBtn.classList.add('active');
  } else {
    // Switch back to rich editor mode
    const htmlTextarea = document.getElementById(editorElement.id + '-html');
    if (htmlTextarea) {
      editorElement.innerHTML = htmlTextarea.value;
      hiddenInput.value = htmlTextarea.value;
      htmlTextarea.style.display = 'none';
    }
    editorElement.style.display = 'block';
    
    // Update the toggle button appearance
    htmlToggleBtn.classList.remove('active');
  }
}

// Load HTML content from file using chrome.runtime.getURL
async function loadTabContent(tabName) {
  console.log(`Loading tab content for: ${tabName}`);
  const url = chrome.runtime.getURL(`src/ui/panel/tabs/${tabName}-tab.html`);
  console.log(`URL for tab content: ${url}`);
  
  try {
    const response = await fetch(url);
    console.log(`Fetch response for ${tabName}:`, response);
    
    if (!response.ok) {
      throw new Error(`Failed to load ${tabName} tab: ${response.status} ${response.statusText}`);
    }
    
    const content = await response.text();
    console.log(`Content loaded for ${tabName}, length: ${content.length} characters`);
    return content;
  } catch (error) {
    console.error(`Error loading ${tabName} tab:`, error);
    return `<div class="error">Failed to load ${tabName} tab content: ${error.message}</div>`;
  }
}

// Initialize tab content loading and switching functionality
async function initTabSwitching() {
  console.log('initTabSwitching called');
  
  const tabs = document.querySelectorAll('.tab');
  console.log(`Found ${tabs.length} tabs`);
  
  const tabContents = document.querySelectorAll('.tab-content');
  console.log(`Found ${tabContents.length} tab content containers`);
  
  // Show loading status
  showStatus('Loading tab content...', 'pending');
  
  try {
    console.log('Loading tab content from HTML files');
    // Load tab content from HTML files
    const emailTabContent = await loadTabContent('email');
    const eventsTabContent = await loadTabContent('events');
    const bulkEmailTabContent = await loadTabContent('bulk-email');
    
    const emailTab = document.getElementById('email-tab');
    const eventsTab = document.getElementById('events-tab');
    const bulkTab = document.getElementById('bulk-tab');
    
    console.log('Email tab element exists:', !!emailTab);
    console.log('Events tab element exists:', !!eventsTab);
    console.log('Bulk tab element exists:', !!bulkTab);
    
    if (emailTab) emailTab.innerHTML = emailTabContent;
    if (eventsTab) eventsTab.innerHTML = eventsTabContent;
    if (bulkTab) bulkTab.innerHTML = bulkEmailTabContent;
    
    // Make sure the active tab is visible
    const activeTab = document.querySelector('.tab.active');
    if (activeTab) {
      const activeTabId = activeTab.getAttribute('data-tab');
      const activeTabContent = document.getElementById(activeTabId + '-tab');
      if (activeTabContent) {
        console.log(`Setting active tab content: ${activeTabId}-tab`);
        // Make sure all tab contents are hidden first
        tabContents.forEach(c => c.style.display = 'none');
        // Then make the active one visible
        activeTabContent.style.display = 'block';
      }
    }
    
    // Clear loading status
    showStatus('');
    
    // Initialize tab switching
    console.log('Setting up tab click handlers');
    tabs.forEach(tab => {
      tab.addEventListener('click', function() {
        console.log(`Tab clicked: ${this.getAttribute('data-tab')}`);
        // Remove active class from all tabs and contents
        tabs.forEach(t => t.classList.remove('active'));
        tabContents.forEach(c => {
          c.classList.remove('active');
          c.style.display = 'none';
        });
        
        // Add active class to clicked tab and corresponding content
        const tabId = this.getAttribute('data-tab');
        this.classList.add('active');
        const activeContent = document.getElementById(tabId + '-tab');
        activeContent.classList.add('active');
        activeContent.style.display = 'block';
        console.log(`Activated tab: ${tabId}, display:`, activeContent.style.display);
        
        // Clear status when switching tabs
        showStatus('');
      });
    });
    
    console.log('Tab switching initialization complete');
    return true;
  } catch (error) {
    console.error('Error loading tab content:', error);
    showStatus('Failed to load tab content. Please refresh the page.', 'error');
    return false;
  }
}

// Export functions
window.CommonUtils = {
  validateEmail,
  showStatus,
  initRichTextEditors,
  initTabSwitching
};
