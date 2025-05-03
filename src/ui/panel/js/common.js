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

// HTML templates for each tab
const tabTemplates = {
  email: `<div class="form-group">
    <label for="to">To:</label>
    <input type="email" id="to" placeholder="recipient@example.com">
  </div>
  <div class="form-group">
    <label for="subject">Subject:</label>
    <input type="text" id="subject" placeholder="Email subject">
  </div>
  <div class="form-group">
    <label for="body">Body (HTML supported):</label>
    <div class="editor-container">
      <div class="editor-toolbar" id="body-toolbar">
        <select id="body-format-select">
          <option value="p">Paragraph</option>
          <option value="h1">Heading 1</option>
          <option value="h2">Heading 2</option>
          <option value="h3">Heading 3</option>
        </select>
        <button type="button" data-command="bold" title="Bold"><strong>B</strong></button>
        <button type="button" data-command="italic" title="Italic"><em>I</em></button>
        <button type="button" data-command="underline" title="Underline"><u>U</u></button>
        <button type="button" data-command="foreColor" data-value="#FF0000" title="Text Color">Color</button>
        <button type="button" data-command="insertOrderedList" title="Numbered List">1.</button>
        <button type="button" data-command="insertUnorderedList" title="Bullet List">•</button>
        <button type="button" data-command="createLink" title="Insert Link">Link</button>
        <button type="button" data-command="justifyLeft" title="Align Left">Left</button>
        <button type="button" data-command="justifyCenter" title="Align Center">Center</button>
        <button type="button" data-command="justifyRight" title="Align Right">Right</button>
        <button type="button" id="body-html-toggle-btn" class="html-toggle-btn" title="Edit HTML directly">HTML</button>
      </div>
      <div id="body-editor" class="rich-editor" contenteditable="true" placeholder="Enter your email content here..."></div>
      <textarea id="body" style="display:none;"></textarea>
    </div>
    <div class="editor-toggle" style="display: none;">
      <input type="checkbox" id="body-html-toggle">
      <label for="body-html-toggle">Edit HTML directly</label>
    </div>
  </div>
  <div class="buttons">
    <button id="sendBtn" class="primary-button">Send Email</button>
    <button id="clearBtn" class="secondary-button">Clear</button>
  </div>`,
  
  events: `<div class="form-group">
    <label for="eventId">Event Title or ID:</label>
    <input type="text" id="eventId" placeholder="Enter event title or ID">
  </div>
  <div class="buttons">
    <button id="getEventBtn" class="primary-button">Get Event Details</button>
  </div>
  <div id="event-result" style="display: none;">
    <h4>Event Details</h4>
    <div class="event-details">
      <div class="detail-row">
        <span class="detail-label">Title:</span>
        <span id="event-title" class="detail-value"></span>
      </div>
      <div class="detail-row">
        <span class="detail-label">Organizer:</span>
        <span id="event-organizer" class="detail-value"></span>
      </div>
      <div class="detail-row">
        <span class="detail-label">Date/Time:</span>
        <span id="event-datetime" class="detail-value"></span>
      </div>
      <div class="detail-row">
        <span class="detail-label">Location:</span>
        <span id="event-location" class="detail-value"></span>
      </div>
    </div>
    <h4>Attendees <span id="attendee-count">(0)</span></h4>
    <div id="attendee-list" class="attendee-list">
      <!-- Attendees will be inserted here -->
    </div>
    <div class="buttons" style="margin-top: 10px;">
      <button id="copyEmailsBtn" class="secondary-button">Copy All Emails</button>
    </div>
  </div>`,
  
  bulk: `<div class="form-group">
    <label for="bulk-subject">Subject ({{placeholders}} supported):</label>
    <input type="text" id="bulk-subject" placeholder="Email subject with optional {{placeholders}}">
  </div>
  <div class="form-group">
    <label for="bulk-body">Body (HTML and {{placeholders}} supported):</label>
    <div class="editor-container">
      <div class="editor-toolbar" id="bulk-body-toolbar">
        <select id="bulk-body-format-select">
          <option value="p">Paragraph</option>
          <option value="h1">Heading 1</option>
          <option value="h2">Heading 2</option>
          <option value="h3">Heading 3</option>
        </select>
        <button type="button" data-command="bold" title="Bold"><strong>B</strong></button>
        <button type="button" data-command="italic" title="Italic"><em>I</em></button>
        <button type="button" data-command="underline" title="Underline"><u>U</u></button>
        <button type="button" data-command="foreColor" data-value="#FF0000" title="Text Color">Color</button>
        <button type="button" data-command="insertOrderedList" title="Numbered List">1.</button>
        <button type="button" data-command="insertUnorderedList" title="Bullet List">•</button>
        <button type="button" data-command="createLink" title="Insert Link">Link</button>
        <button type="button" data-command="justifyLeft" title="Align Left">Left</button>
        <button type="button" data-command="justifyCenter" title="Align Center">Center</button>
        <button type="button" data-command="justifyRight" title="Align Right">Right</button>
        <button type="button" id="bulk-body-html-toggle-btn" class="html-toggle-btn" title="Edit HTML directly">HTML</button>
      </div>
      <div id="bulk-body-editor" class="rich-editor" contenteditable="true" placeholder="Enter your email content with placeholders like {{name}}, {{company}}, etc."></div>
      <textarea id="bulk-body" style="display:none;"></textarea>
    </div>
    <div class="editor-toggle" style="display: none;">
      <input type="checkbox" id="bulk-body-html-toggle">
      <label for="bulk-body-html-toggle">Edit HTML directly</label>
    </div>
  </div>
    <div class="buttons">
    <button id="previewBtn" class="secondary-button">Preview Data</button>
    <button id="sendBulkBtn" class="primary-button">Send Bulk Emails</button>
    <button id="clearBulkBtn" class="secondary-button">Clear</button>
  </div>
  <div class="form-group">
    <div id="preview-container" style="display: none;">
      <h4>Recipient Preview</h4>
      <div class="preview-info">
        <span id="recipient-count">0</span> recipients found. 
        <span id="placeholder-info"></span>
      </div>
      <div class="select-all-container">
        <input type="checkbox" id="select-all-checkbox" checked>
        <label for="select-all-checkbox">Select/Deselect All</label>
      </div>
      <div class="preview-table-container">
        <table id="preview-table" class="preview-table">
          <!-- Editable table will be inserted here -->
        </table>
      </div>
      <div class="table-actions">
        <button id="addRowBtn" class="secondary-button">Add Row</button>
        <button id="removeSelectedBtn" class="secondary-button">Remove Selected</button>
        <button id="applyChangesBtn" class="secondary-button">Apply Changes</button>
      </div>
    </div>
  </div>

  <div class="form-group">
    <label for="bulk-recipients">Recipients Data (CSV/Tab/Semicolon separated):</label>
    <textarea id="bulk-recipients" rows="6" placeholder="Paste CSV, tab, or semicolon separated data. First row should contain headers including 'email' field."></textarea>
    <div class="recipients-data-container" style="margin-top: 8px;">
      <label>Data Format:</label>
      <div class="radio-group">
        <input type="radio" id="format-auto" name="data-format" value="auto" checked>
        <label for="format-auto">Auto-detect</label>
        
        <input type="radio" id="format-comma" name="data-format" value="comma">
        <label for="format-comma">Comma (,)</label>
        
        <input type="radio" id="format-semicolon" name="data-format" value="semicolon">
        <label for="format-semicolon">Semicolon (;)</label>
        
        <input type="radio" id="format-tab" name="data-format" value="tab">
        <label for="format-tab">Tab</label>
      </div>
    </div>
    <div class="data-management" style="margin-top: 15px;">
      <div class="data-management-buttons">
        <button id="saveDataSetBtn" class="secondary-button">Save Data Set</button>
        <button id="manageDataSetsBtn" class="secondary-button">Manage Data Sets</button>
      </div>
    </div>
  </div>
`
};

// Initialize tab content loading and switching functionality
function initTabSwitching() {
  console.log('initTabSwitching called');
  
  const tabs = document.querySelectorAll('.tab');
  console.log(`Found ${tabs.length} tabs`);
  
  const tabContents = document.querySelectorAll('.tab-content');
  console.log(`Found ${tabContents.length} tab content containers`);
  
  try {
    console.log('Setting tab content from templates');
    // Load tab content from embedded templates
    const emailTab = document.getElementById('email-tab');
    const eventsTab = document.getElementById('events-tab');
    const bulkTab = document.getElementById('bulk-tab');
    
    console.log('Email tab element exists:', !!emailTab);
    console.log('Events tab element exists:', !!eventsTab);
    console.log('Bulk tab element exists:', !!bulkTab);
    
    if (emailTab) emailTab.innerHTML = tabTemplates.email;
    if (eventsTab) eventsTab.innerHTML = tabTemplates.events;
    if (bulkTab) bulkTab.innerHTML = tabTemplates.bulk;
    
    // Initialize tab switching
    console.log('Setting up tab click handlers');
    tabs.forEach(tab => {
      tab.addEventListener('click', function() {
        console.log(`Tab clicked: ${this.getAttribute('data-tab')}`);
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
