document.addEventListener('DOMContentLoaded', function() {
  // Email tab elements
  const toInput = document.getElementById('to');
  const subjectInput = document.getElementById('subject');
  const bodyInput = document.getElementById('body');
  const bodyEditor = document.getElementById('body-editor');
  const bodyToolbar = document.getElementById('body-toolbar');
  const bodyFormatSelect = document.getElementById('body-format-select');
  const bodyHtmlToggle = document.getElementById('body-html-toggle');
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
  
  // Bulk Email tab elements
  const bulkSubjectInput = document.getElementById('bulk-subject');
  const bulkBodyInput = document.getElementById('bulk-body');
  const bulkBodyEditor = document.getElementById('bulk-body-editor');
  const bulkBodyToolbar = document.getElementById('bulk-body-toolbar');
  const bulkBodyFormatSelect = document.getElementById('bulk-body-format-select');
  const bulkBodyHtmlToggle = document.getElementById('bulk-body-html-toggle');
  const bulkRecipientsInput = document.getElementById('bulk-recipients');
  const previewBtn = document.getElementById('previewBtn');
  const sendBulkBtn = document.getElementById('sendBulkBtn');
  const clearBulkBtn = document.getElementById('clearBulkBtn');
  const previewContainer = document.getElementById('preview-container');
  const previewTable = document.getElementById('preview-table');
  const recipientCountSpan = document.getElementById('recipient-count');
  const placeholderInfoSpan = document.getElementById('placeholder-info');
  const formatRadios = document.querySelectorAll('input[name="data-format"]');
  const selectAllCheckbox = document.getElementById('select-all-checkbox');
  const addRowBtn = document.getElementById('addRowBtn');
  const removeSelectedBtn = document.getElementById('removeSelectedBtn');
  const applyChangesBtn = document.getElementById('applyChangesBtn');
  
  // Store the parsed data for editing
  let tableData = {
    headers: [],
    records: []
  };
  
  // Tab navigation elements
  const tabs = document.querySelectorAll('.tab');
  const tabContents = document.querySelectorAll('.tab-content');
  
  // Status element
  const statusDiv = document.getElementById('status');

  // Initialize rich text editors after all DOM elements are defined
  initRichTextEditors();

  // Load any saved draft, event ID, and bulk email data
  chrome.storage.local.get([
    'draftTo', 'draftSubject', 'draftBody', 'lastEventId',
    'bulkSubject', 'bulkBody', 'bulkRecipients', 'bulkFormat'
  ], function(result) {
    if (result.draftTo) toInput.value = result.draftTo;
    if (result.draftSubject) subjectInput.value = result.draftSubject;
    if (result.draftBody) bodyInput.value = result.draftBody;
    if (result.lastEventId) eventIdInput.value = result.lastEventId;
    if (result.bulkSubject) bulkSubjectInput.value = result.bulkSubject;
    if (result.bulkBody) bulkBodyInput.value = result.bulkBody;
    if (result.bulkRecipients) bulkRecipientsInput.value = result.bulkRecipients;
    if (result.bulkFormat) {
      document.getElementById('format-' + result.bulkFormat).checked = true;
    }
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
      
      // Update rich text editors when switching tabs
      if (tabId === 'email' || tabId === 'bulk') {
        // No special action needed for contenteditable editors
      }
    });
  });

  // Save draft as user types
  function saveDraft() {
    // Get content from the editor
    const bodyContent = bodyInput.value || bodyEditor.innerHTML;
    
    chrome.storage.local.set({
      draftTo: toInput.value,
      draftSubject: subjectInput.value,
      draftBody: bodyContent
    });
  }
  
  // Save event ID when user types
  function saveEventId() {
    chrome.storage.local.set({
      lastEventId: eventIdInput.value
    });
  }
  
  // Save bulk email data when user types
  function saveBulkData() {
    let selectedFormat = 'auto';
    formatRadios.forEach(radio => {
      if (radio.checked) {
        selectedFormat = radio.value;
      }
    });
    
    // Get content from the editor
    const bodyContent = bulkBodyInput.value || bulkBodyEditor.innerHTML;
    
    chrome.storage.local.set({
      bulkSubject: bulkSubjectInput.value,
      bulkBody: bodyContent,
      bulkRecipients: bulkRecipientsInput.value,
      bulkFormat: selectedFormat
    });
  }

  toInput.addEventListener('input', saveDraft);
  subjectInput.addEventListener('input', saveDraft);
  bodyInput.addEventListener('input', saveDraft);
  eventIdInput.addEventListener('input', saveEventId);
  
  // Save bulk email data as user types
  bulkSubjectInput.addEventListener('input', saveBulkData);
  bulkBodyInput.addEventListener('input', saveBulkData);
  bulkRecipientsInput.addEventListener('input', saveBulkData);
  formatRadios.forEach(radio => {
    radio.addEventListener('change', saveBulkData);
  });

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
    
    // Clear editor content
    bodyEditor.innerHTML = '';
    bodyInput.value = '';
    
    chrome.storage.local.remove(['draftTo', 'draftSubject', 'draftBody']);
    statusDiv.textContent = '';
    statusDiv.className = 'status';
  }
  
  // Clear bulk form
  clearBulkBtn.addEventListener('click', clearBulkForm);
  
  function clearBulkForm() {
    bulkSubjectInput.value = '';
    
    // Clear editor content
    bulkBodyEditor.innerHTML = '';
    bulkBodyInput.value = '';
    
    bulkRecipientsInput.value = '';
    document.getElementById('format-auto').checked = true;
    chrome.storage.local.remove(['bulkSubject', 'bulkBody', 'bulkRecipients', 'bulkFormat']);
    previewContainer.style.display = 'none';
    statusDiv.textContent = '';
    statusDiv.className = 'status';
    
    // Clear the table data
    tableData = {
      headers: [],
      records: []
    };
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
  
  // Initialize rich text editors
  function initRichTextEditors() {
    // Get HTML toggle buttons from toolbar
    const bodyHtmlToggleBtn = document.getElementById('body-html-toggle-btn');
    const bulkBodyHtmlToggleBtn = document.getElementById('bulk-body-html-toggle-btn');
    
    // Set up the regular email editor
    setupEditor(bodyEditor, bodyToolbar, bodyFormatSelect, bodyInput, bodyHtmlToggle, bodyHtmlToggleBtn, saveDraft);
    
    // Set up the bulk email editor
    setupEditor(bulkBodyEditor, bulkBodyToolbar, bulkBodyFormatSelect, bulkBodyInput, bulkBodyHtmlToggle, bulkBodyHtmlToggleBtn, saveBulkData);
    
    // Load saved content
    chrome.storage.local.get(['draftBody', 'bulkBody'], function(result) {
      if (result.draftBody) {
        bodyEditor.innerHTML = result.draftBody;
        bodyInput.value = result.draftBody;
      }
      
      if (result.bulkBody) {
        bulkBodyEditor.innerHTML = result.bulkBody;
        bulkBodyInput.value = result.bulkBody;
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
  
  // Parse CSV/TSV/Semicolon-separated data
  function parseData(data, format) {
    if (!data.trim()) {
      return { headers: [], records: [] };
    }
    
    // Determine the delimiter based on format or auto-detect
    let delimiter;
    if (format === 'comma') {
      delimiter = ',';
    } else if (format === 'semicolon') {
      delimiter = ';';
    } else if (format === 'tab') {
      delimiter = '\t';
    } else {
      // Auto-detect delimiter
      const firstLine = data.trim().split('\n')[0];
      if (firstLine.includes(';')) {
        delimiter = ';';
      } else if (firstLine.includes('\t')) {
        delimiter = '\t';
      } else {
        delimiter = ',';
      }
    }
    
    // Split into lines and remove empty lines
    const lines = data.trim().split('\n').filter(line => line.trim());
    if (lines.length === 0) {
      return { headers: [], records: [] };
    }
    
    // Parse headers (first line)
    const headers = lines[0].split(delimiter).map(header => {
      // Remove quotes if present
      return header.trim().replace(/^"|"$/g, '');
    });
    
    // Parse records
    const records = [];
    for (let i = 1; i < lines.length; i++) {
      const values = lines[i].split(delimiter).map(value => {
        // Remove quotes if present
        return value.trim().replace(/^"|"$/g, '');
      });
      
      // Create record object with header keys
      const record = {};
      headers.forEach((header, index) => {
        record[header] = values[index] || '';
      });
      
      records.push(record);
    }
    
    return { headers, records };
  }
  
  // Extract placeholders from a template string
  function extractPlaceholders(template) {
    const placeholderRegex = /\{\{([^{}]+)\}\}/g;
    const placeholders = new Set();
    let match;
    
    while ((match = placeholderRegex.exec(template)) !== null) {
      placeholders.add(match[1]);
    }
    
    return Array.from(placeholders);
  }
  
  // Replace placeholders in a template with values from a record
  function replacePlaceholders(template, record) {
    return template.replace(/\{\{([^{}]+)\}\}/g, (match, placeholder) => {
      return record[placeholder] || match;
    });
  }
  
  // Preview button click handler
  previewBtn.addEventListener('click', function() {
    const recipientsData = bulkRecipientsInput.value.trim();
    if (!recipientsData) {
      showStatus('Please paste recipient data', 'error');
      return;
    }
    
    // Get selected format
    let selectedFormat = 'auto';
    formatRadios.forEach(radio => {
      if (radio.checked) {
        selectedFormat = radio.value;
      }
    });
    
    // Parse the data
    const parsedData = parseData(recipientsData, selectedFormat);
    if (parsedData.records.length === 0) {
      showStatus('No valid data found or invalid format', 'error');
      return;
    }
    
    // Store the parsed data for editing
    tableData = parsedData;
    
    // Check if email field exists
    const emailField = tableData.headers.find(header => 
      ['email', 'emailaddress', 'e-mail', 'mail'].includes(header.toLowerCase())
    );
    
    if (!emailField) {
      showStatus('No email field found in the data. Please include a column named "email" or similar.', 'error');
      return;
    }
    
    // Extract placeholders from subject and body
    const subjectPlaceholders = extractPlaceholders(bulkSubjectInput.value);
    const bodyPlaceholders = extractPlaceholders(bulkBodyInput.value);
    const allPlaceholders = [...new Set([...subjectPlaceholders, ...bodyPlaceholders])];
    
    // Update UI
    recipientCountSpan.textContent = tableData.records.length;
    
    if (allPlaceholders.length > 0) {
      placeholderInfoSpan.innerHTML = `Found placeholders: ${allPlaceholders.map(p => 
        `<span class="placeholder-highlight">{{${p}}}</span>`).join(', ')}`;
    } else {
      placeholderInfoSpan.textContent = 'No placeholders found in subject or body.';
    }
    
    // Generate editable table
    renderEditableTable();
    
    previewContainer.style.display = 'block';
    showStatus(`Data loaded. ${tableData.records.length} recipients ready for editing.`, 'success');
  });
  
  // Render the editable table with the current data
  function renderEditableTable() {
    // Clear the table
    previewTable.innerHTML = '';
    
    // Create header row
    const headerRow = document.createElement('tr');
    
    // Add checkbox column header
    const checkboxHeader = document.createElement('th');
    checkboxHeader.style.width = '30px';
    checkboxHeader.textContent = '';
    headerRow.appendChild(checkboxHeader);
    
    // Add editable headers
    tableData.headers.forEach((header, index) => {
      const th = document.createElement('th');
      const input = document.createElement('input');
      input.type = 'text';
      input.value = header;
      input.dataset.index = index;
      input.className = 'header-input';
      input.addEventListener('change', function() {
        // Store the old header name before updating it
        const oldHeader = tableData.headers[index];
        const newHeader = this.value;
        
        // Update the header in the array
        tableData.headers[index] = newHeader;
        
        // Update all records to use the new header name
        if (oldHeader !== newHeader) {
          tableData.records.forEach(record => {
            // Only transfer the value if the old header exists in the record
            if (record.hasOwnProperty(oldHeader)) {
              // Copy value to new header name
              record[newHeader] = record[oldHeader];
              // Remove old header property if different
              if (oldHeader !== newHeader) {
                delete record[oldHeader];
              }
            }
          });
        }
      });
      th.appendChild(input);
      headerRow.appendChild(th);
    });
    
    previewTable.appendChild(headerRow);
    
    // Create data rows
    tableData.records.forEach((record, rowIndex) => {
      const row = document.createElement('tr');
      row.dataset.index = rowIndex;
      
      // Add checkbox cell
      const checkboxCell = document.createElement('td');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = true;
      checkbox.className = 'row-checkbox';
      checkbox.dataset.index = rowIndex;
      checkboxCell.appendChild(checkbox);
      row.appendChild(checkboxCell);
      
      // Add editable data cells
      tableData.headers.forEach((header, colIndex) => {
        const td = document.createElement('td');
        const input = document.createElement('input');
        input.type = 'text';
        input.value = record[header] || '';
        input.dataset.row = rowIndex;
        input.dataset.header = header;
        input.addEventListener('change', function() {
          tableData.records[rowIndex][header] = this.value;
        });
        td.appendChild(input);
        row.appendChild(td);
      });
      
      previewTable.appendChild(row);
    });
  }
  
  // Select/deselect all checkboxes
  selectAllCheckbox.addEventListener('change', function() {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    checkboxes.forEach(checkbox => {
      checkbox.checked = selectAllCheckbox.checked;
    });
  });
  
  // Add a new row
  addRowBtn.addEventListener('click', function() {
    // Create a new empty record
    const newRecord = {};
    tableData.headers.forEach(header => {
      newRecord[header] = '';
    });
    
    // Add to the records array
    tableData.records.push(newRecord);
    
    // Update the UI
    renderEditableTable();
    recipientCountSpan.textContent = tableData.records.length;
    showStatus('New row added', 'success');
  });
  
  // Remove selected rows
  removeSelectedBtn.addEventListener('click', function() {
    const checkboxes = document.querySelectorAll('.row-checkbox:not(:checked)');
    if (checkboxes.length === 0) {
      showStatus('No rows selected for removal', 'error');
      return;
    }
    
    // Get indices of rows to keep (those that are checked)
    const indicesToRemove = Array.from(checkboxes).map(checkbox => 
      parseInt(checkbox.dataset.index)
    );
    
    // Filter out the records to remove
    tableData.records = tableData.records.filter((_, index) => 
      !indicesToRemove.includes(index)
    );
    
    // Update the UI
    renderEditableTable();
    recipientCountSpan.textContent = tableData.records.length;
    showStatus(`${indicesToRemove.length} row(s) removed`, 'success');
  });
  
  // Apply changes button
  applyChangesBtn.addEventListener('click', function() {
    // Get the current header inputs to check for changes
    const headerInputs = document.querySelectorAll('.header-input');
    const newHeaders = Array.from(headerInputs).map(input => input.value);
    
    // Check for header name changes and update record keys
    const headerChanges = {};
    tableData.headers.forEach((oldHeader, index) => {
      const newHeader = newHeaders[index];
      if (oldHeader !== newHeader) {
        headerChanges[oldHeader] = newHeader;
      }
    });
    
    // Update record keys if headers changed
    if (Object.keys(headerChanges).length > 0) {
      tableData.records.forEach(record => {
        Object.keys(headerChanges).forEach(oldHeader => {
          const newHeader = headerChanges[oldHeader];
          // Copy the value from old header to new header
          if (record.hasOwnProperty(oldHeader)) {
            record[newHeader] = record[oldHeader];
            // Only delete the old key if it's different from the new one
            if (oldHeader !== newHeader) {
              delete record[oldHeader];
            }
          }
        });
      });
      
      // Update the headers array
      tableData.headers = newHeaders;
    }
    
    // Update the CSV data in the textarea based on the edited table
    let selectedFormat = 'auto';
    formatRadios.forEach(radio => {
      if (radio.checked) {
        selectedFormat = radio.value;
      }
    });
    
    let delimiter;
    if (selectedFormat === 'comma') {
      delimiter = ',';
    } else if (selectedFormat === 'semicolon') {
      delimiter = ';';
    } else if (selectedFormat === 'tab') {
      delimiter = '\t';
    } else {
      delimiter = ',';
    }
    
    // Generate CSV string
    let csvContent = tableData.headers.join(delimiter) + '\n';
    
    // Only include checked rows
    const checkboxes = document.querySelectorAll('.row-checkbox');
    const checkedIndices = Array.from(checkboxes)
      .filter(checkbox => checkbox.checked)
      .map(checkbox => parseInt(checkbox.dataset.index));
    
    // Add rows for checked records
    checkedIndices.forEach(index => {
      const record = tableData.records[index];
      const row = tableData.headers.map(header => record[header] || '').join(delimiter);
      csvContent += row + '\n';
    });
    
    // Update the textarea
    bulkRecipientsInput.value = csvContent;
    
    showStatus('Changes applied to data', 'success');
  });
  
  // Send bulk emails button click handler
  sendBulkBtn.addEventListener('click', function() {
    const subject = bulkSubjectInput.value.trim();
    const body = bulkBodyInput.value.trim();
    
    // Validate inputs
    if (!subject) {
      showStatus('Please enter a subject', 'error');
      return;
    }
    
    if (!body) {
      showStatus('Please enter a message body', 'error');
      return;
    }
    
    // Check if we have data loaded
    if (tableData.records.length === 0) {
      showStatus('Please load and preview recipient data first', 'error');
      return;
    }
    
    // Get only the selected records
    const checkboxes = document.querySelectorAll('.row-checkbox');
    const selectedRecords = [];
    checkboxes.forEach((checkbox, index) => {
      if (checkbox.checked) {
        selectedRecords.push(tableData.records[index]);
      }
    });
    
    if (selectedRecords.length === 0) {
      showStatus('No recipients selected. Please select at least one recipient.', 'error');
      return;
    }
    
    // Check if email field exists
    const emailField = tableData.headers.find(header => 
      ['email', 'emailaddress', 'e-mail', 'mail'].includes(header.toLowerCase())
    );
    
    if (!emailField) {
      showStatus('No email field found in the data. Please include a column named "email" or similar.', 'error');
      return;
    }
    
    // First check if current tab is Outlook
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      const currentTab = tabs[0];
      const isOutlook = currentTab && currentTab.url && 
                     (currentTab.url.includes('outlook.office.com') || 
                      currentTab.url.includes('outlook.live.com'));
      
      if (isOutlook) {
        // Current tab is Outlook, send directly
        processBulkEmails(currentTab.id, selectedRecords, emailField, subject, body);
      } else {
        // Current tab is not Outlook, try to find or create an Outlook tab
        chrome.runtime.sendMessage({ action: 'activateOutlookTab' }, function(result) {
          if (result && result.success) {
            // Give the tab time to load if it's new
            const delay = result.existing ? 500 : 3000;
            showStatus(`${result.existing ? 'Switching to' : 'Opening'} Outlook tab...`, 'pending');
            
            setTimeout(() => {
              processBulkEmails(result.tabId, selectedRecords, emailField, subject, body);
            }, delay);
          } else {
            showStatus('Failed to open Outlook tab: ' + (result?.error || 'Unknown error'), 'error');
          }
        });
      }
    });
  });
  
  // Process and send bulk emails
  function processBulkEmails(tabId, records, emailField, subjectTemplate, bodyTemplate) {
    showStatus(`Preparing to send ${records.length} emails...`, 'pending');
    
    // Process one email at a time
    let currentIndex = 0;
    
    function sendNextEmail() {
      if (currentIndex >= records.length) {
        showStatus(`All ${records.length} emails have been processed!`, 'success');
        return;
      }
      
      const record = records[currentIndex];
      const emailAddress = record[emailField];
      
      if (!emailAddress || !validateEmail(emailAddress)) {
        showStatus(`Skipping record ${currentIndex + 1}: Invalid email address`, 'error');
        currentIndex++;
        setTimeout(sendNextEmail, 500);
        return;
      }
      
      // Replace placeholders in subject and body
      const subject = replacePlaceholders(subjectTemplate, record);
      const body = replacePlaceholders(bodyTemplate, record);
      
      // Prepare email data
      const emailData = {
        to: emailAddress,
        subject: subject,
        body: body
      };
      
      showStatus(`Sending email ${currentIndex + 1} of ${records.length} to ${emailAddress}...`, 'pending');
      
      // Send the email
      chrome.tabs.sendMessage(tabId, {
        action: 'sendEmail',
        data: emailData
      }, function(response) {
        if (chrome.runtime.lastError) {
          showStatus(`Error sending to ${emailAddress}: ${chrome.runtime.lastError.message}`, 'error');
          currentIndex++;
          setTimeout(sendNextEmail, 2000); // Wait a bit longer after an error
          return;
        }
        
        if (response && response.success) {
          showStatus(`Email ${currentIndex + 1} sent to ${emailAddress}`, 'success');
        } else {
          showStatus(`Failed to send to ${emailAddress}: ${response?.error || 'Unknown error'}`, 'error');
        }
        
        currentIndex++;
        setTimeout(sendNextEmail, 2000); // Wait between emails to avoid overwhelming Outlook
      });
    }
    
    // Start sending emails
    sendNextEmail();
  }
});
