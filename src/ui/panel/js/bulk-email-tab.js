// bulk-email-tab.js - Bulk email tab functionality

// Initialize the bulk email tab
function initBulkEmailTab() {
  // Bulk Email tab elements
  const bulkSubjectInput = document.getElementById('bulk-subject');
  const bulkBodyInput = document.getElementById('bulk-body');
  const bulkBodyEditor = document.getElementById('bulk-body-editor');
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
  const saveDataSetBtn = document.getElementById('saveDataSetBtn');
  const manageDataSetsBtn = document.getElementById('manageDataSetsBtn');
  
  // Store the parsed data for editing
  let tableData = {
    headers: [],
    records: []
  };
  
  // Load any saved bulk email data
  chrome.storage.local.get([
    'bulkSubject', 'bulkBody', 'bulkRecipients', 'bulkFormat'
  ], function(result) {
    if (result.bulkSubject) bulkSubjectInput.value = result.bulkSubject;
    // Note: body content is loaded in common.js
    if (result.bulkRecipients) bulkRecipientsInput.value = result.bulkRecipients;
    if (result.bulkFormat) {
      document.getElementById('format-' + result.bulkFormat).checked = true;
    }
  });
  
  // Save bulk email data as user types
  bulkSubjectInput.addEventListener('input', saveBulkData);
  bulkRecipientsInput.addEventListener('input', saveBulkData);
  formatRadios.forEach(radio => {
    radio.addEventListener('change', saveBulkData);
  });
  
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
  
  // Preview button click handler
  previewBtn.addEventListener('click', function() {
    const recipientsData = bulkRecipientsInput.value.trim();
    if (!recipientsData) {
      CommonUtils.showStatus('Please paste recipient data', 'error');
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
      CommonUtils.showStatus('No valid data found or invalid format', 'error');
      return;
    }
    
    // Store the parsed data for editing
    tableData = parsedData;
    
    // Check if email field exists
    const emailField = tableData.headers.find(header => 
      ['email', 'emailaddress', 'e-mail', 'mail'].includes(header.toLowerCase())
    );
    
    if (!emailField) {
      CommonUtils.showStatus('No email field found in the data. Please include a column named "email" or similar.', 'error');
      return;
    }
    
    // Extract placeholders from subject and body
    const subjectPlaceholders = extractPlaceholders(bulkSubjectInput.value);
    const bodyPlaceholders = extractPlaceholders(bulkBodyInput.value || bulkBodyEditor.innerHTML);
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
    CommonUtils.showStatus(`Data loaded. ${tableData.records.length} recipients ready for editing.`, 'success');
  });
  
  // Track current sort state
  let currentSort = {
    column: null,
    direction: 'asc'
  };
  
  // Sort the table data by column
  function sortTableData(columnIndex) {
    const columnName = tableData.headers[columnIndex];
    
    // If already sorting by this column, toggle direction
    if (currentSort.column === columnIndex) {
      currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
      currentSort.column = columnIndex;
      currentSort.direction = 'asc';
    }
    
    // Sort the records
    tableData.records.sort((a, b) => {
      const valueA = (a[columnName] || '').toString().toLowerCase();
      const valueB = (b[columnName] || '').toString().toLowerCase();
      
      // Check if values are numeric
      const numA = parseFloat(valueA);
      const numB = parseFloat(valueB);
      
      if (!isNaN(numA) && !isNaN(numB)) {
        // Numeric sort
        return currentSort.direction === 'asc' ? numA - numB : numB - numA;
      } else {
        // String sort
        if (valueA < valueB) return currentSort.direction === 'asc' ? -1 : 1;
        if (valueA > valueB) return currentSort.direction === 'asc' ? 1 : -1;
        return 0;
      }
    });
    
    // Re-render the table with sorted data
    renderEditableTable();
  }
  
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
    
    // Add editable headers with sort functionality
    tableData.headers.forEach((header, index) => {
      const th = document.createElement('th');
      
      // Create header container to hold input and sort indicator
      const headerContainer = document.createElement('div');
      headerContainer.className = 'header-container';
      headerContainer.style.display = 'flex';
      headerContainer.style.alignItems = 'center';
      
      // Create input for header name
      const input = document.createElement('input');
      input.type = 'text';
      input.value = header;
      input.dataset.index = index;
      input.className = 'header-input';
      input.style.flex = '1';
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
      
      // Create sort indicator
      const sortIndicator = document.createElement('span');
      sortIndicator.className = 'sort-indicator';
      sortIndicator.style.marginLeft = '5px';
      sortIndicator.style.cursor = 'pointer';
      sortIndicator.style.userSelect = 'none';
      
      // Set sort indicator text based on current sort state
      if (currentSort.column === index) {
        sortIndicator.textContent = currentSort.direction === 'asc' ? '^' : 'v';
        sortIndicator.style.color = '#0078d4';
      } else {
        sortIndicator.textContent = '^v';
        sortIndicator.style.color = '#666';
      }
      
      // Add click handler for sorting
      sortIndicator.addEventListener('click', function() {
        sortTableData(index);
      });
      
      // Add elements to container
      headerContainer.appendChild(input);
      headerContainer.appendChild(sortIndicator);
      th.appendChild(headerContainer);
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
      
      // Add preview button cell
      const previewCell = document.createElement('td');
      const previewButton = document.createElement('button');
      previewButton.textContent = 'Preview';
      previewButton.className = 'preview-button';
      previewButton.title = 'Preview email with this recipient\'s data';
      previewButton.dataset.index = rowIndex;
      previewButton.addEventListener('click', function() {
        showEmailPreview(record);
      });
      previewCell.appendChild(previewButton);
      row.appendChild(previewCell);
      
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
    CommonUtils.showStatus('New row added', 'success');
  });
  
  // Remove selected rows
  removeSelectedBtn.addEventListener('click', function() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    if (checkboxes.length === 0) {
      CommonUtils.showStatus('No rows selected for removal', 'error');
      return;
    }
    
    // Get indices of rows to remove (those that are checked)
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
    CommonUtils.showStatus(`${indicesToRemove.length} row(s) removed`, 'success');
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
    
    CommonUtils.showStatus('Changes applied to data', 'success');
  });
  
  // Send bulk emails button click handler
  sendBulkBtn.addEventListener('click', function() {
    const subject = bulkSubjectInput.value.trim();
    const body = bulkBodyInput.value.trim() || bulkBodyEditor.innerHTML;
    
    // Validate inputs
    if (!subject) {
      CommonUtils.showStatus('Please enter a subject', 'error');
      return;
    }
    
    if (!body) {
      CommonUtils.showStatus('Please enter a message body', 'error');
      return;
    }
    
    // Check if we have data loaded
    if (tableData.records.length === 0) {
      CommonUtils.showStatus('Please load and preview recipient data first', 'error');
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
      CommonUtils.showStatus('No recipients selected. Please select at least one recipient.', 'error');
      return;
    }
    
    // Check if email field exists
    const emailField = tableData.headers.find(header => 
      ['email', 'emailaddress', 'e-mail', 'mail'].includes(header.toLowerCase())
    );
    
    if (!emailField) {
      CommonUtils.showStatus('No email field found in the data. Please include a column named "email" or similar.', 'error');
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
            CommonUtils.showStatus(`${result.existing ? 'Switching to' : 'Opening'} Outlook tab...`, 'pending');
            
            setTimeout(() => {
              processBulkEmails(result.tabId, selectedRecords, emailField, subject, body);
            }, delay);
          } else {
            CommonUtils.showStatus('Failed to open Outlook tab: ' + (result?.error || 'Unknown error'), 'error');
          }
        });
      }
    });
  });
  
  // Process and send bulk emails
  function processBulkEmails(tabId, records, emailField, subjectTemplate, bodyTemplate) {
    CommonUtils.showStatus(`Preparing to send ${records.length} emails...`, 'pending');
    
    // Process one email at a time
    let currentIndex = 0;
    
    function sendNextEmail() {
      if (currentIndex >= records.length) {
        CommonUtils.showStatus(`All ${records.length} emails have been processed!`, 'success');
        return;
      }
      
      const record = records[currentIndex];
      const emailAddress = record[emailField];
      
      if (!emailAddress || !CommonUtils.validateEmail(emailAddress)) {
        CommonUtils.showStatus(`Skipping record ${currentIndex + 1}: Invalid email address`, 'error');
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
      
      CommonUtils.showStatus(`Sending email ${currentIndex + 1} of ${records.length} to ${emailAddress}...`, 'pending');
      
      // Send the email
      chrome.tabs.sendMessage(tabId, {
        action: 'sendEmail',
        data: emailData
      }, function(response) {
        if (chrome.runtime.lastError) {
          CommonUtils.showStatus(`Error sending to ${emailAddress}: ${chrome.runtime.lastError.message}`, 'error');
          currentIndex++;
          setTimeout(sendNextEmail, 2000); // Wait a bit longer after an error
          return;
        }
        
        if (response && response.success) {
          CommonUtils.showStatus(`Email ${currentIndex + 1} sent to ${emailAddress}`, 'success');
        } else {
          CommonUtils.showStatus(`Failed to send to ${emailAddress}: ${response?.error || 'Unknown error'}`, 'error');
        }
        
        currentIndex++;
        setTimeout(sendNextEmail, 2000); // Wait between emails to avoid overwhelming Outlook
      });
    }
    
    // Start sending emails
    sendNextEmail();
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
    document.getElementById('status').textContent = '';
    document.getElementById('status').className = 'status';
    
    // Clear the table data
    tableData = {
      headers: [],
      records: []
    };
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
  
  // Extract placeholders from text
  function extractPlaceholders(text) {
    const placeholderRegex = /{{([^{}]+)}}/g;
    const placeholders = [];
    let match;
    
    while ((match = placeholderRegex.exec(text)) !== null) {
      placeholders.push(match[1]);
    }
    
    return [...new Set(placeholders)]; // Remove duplicates
  }
  
  // Replace placeholders in text with values from record
  function replacePlaceholders(text, record) {
    if (!text) return '';
    
    return text.replace(/{{([^{}]+)}}/g, (match, placeholder) => {
      // Check if the placeholder exists in the record
      return record[placeholder] !== undefined ? record[placeholder] : match;
    });
  }
  
  // Show email preview with placeholders replaced for a specific recipient
  function showEmailPreview(record) {
    // Get current subject and body
    const subject = bulkSubjectInput.value;
    const bodyContent = bulkBodyInput.value || bulkBodyEditor.innerHTML;
    
    // Replace placeholders
    const personalizedSubject = replacePlaceholders(subject, record);
    const personalizedBody = replacePlaceholders(bodyContent, record);
    
    // Create or get preview modal
    let previewModal = document.getElementById('email-preview-modal');
    if (!previewModal) {
      previewModal = document.createElement('div');
      previewModal.id = 'email-preview-modal';
      previewModal.className = 'modal';
      previewModal.innerHTML = `
        <div class="modal-content">
          <div class="modal-header">
            <span class="close-modal">&times;</span>
            <h3>Email Preview</h3>
          </div>
          <div class="modal-body">
            <div class="preview-recipient"></div>
            <div class="preview-subject"></div>
            <div class="preview-body"></div>
          </div>
        </div>
      `;
      document.body.appendChild(previewModal);
      
      // Add close button functionality
      const closeButton = previewModal.querySelector('.close-modal');
      closeButton.addEventListener('click', function() {
        previewModal.style.display = 'none';
      });
      
      // Close when clicking outside the modal
      window.addEventListener('click', function(event) {
        if (event.target === previewModal) {
          previewModal.style.display = 'none';
        }
      });
    }
    
    // Update preview content
    const recipientInfo = previewModal.querySelector('.preview-recipient');
    const subjectPreview = previewModal.querySelector('.preview-subject');
    const bodyPreview = previewModal.querySelector('.preview-body');
    
    // Find email field
    const emailField = tableData.headers.find(header => 
      ['email', 'emailaddress', 'e-mail', 'mail'].includes(header.toLowerCase())
    ) || 'email';
    
    recipientInfo.innerHTML = `<strong>To:</strong> ${record[emailField] || 'No email address'}`;
    subjectPreview.innerHTML = `<strong>Subject:</strong> ${personalizedSubject}`;
    bodyPreview.innerHTML = `<strong>Body:</strong><div class="email-body-preview">${personalizedBody}</div>`;
    
    // Show the modal
    previewModal.style.display = 'block';
  }
  
  // Save Data Set button click handler
  saveDataSetBtn.addEventListener('click', function() {
    saveRecipientDataSet();
  });
  
  // Manage Data Sets button click handler
  manageDataSetsBtn.addEventListener('click', function() {
    showDataSetsManager();
  });
  
  // Save current recipient data set with name and description
  function saveRecipientDataSet() {
    // Create modal for saving data set
    let saveModal = document.getElementById('save-dataset-modal');
    if (!saveModal) {
      saveModal = document.createElement('div');
      saveModal.id = 'save-dataset-modal';
      saveModal.className = 'modal';
      saveModal.innerHTML = `
        <div class="modal-content">
          <div class="modal-header">
            <span class="close-modal">&times;</span>
            <h3>Save Recipient Data Set</h3>
          </div>
          <div class="modal-body">
            <div class="form-group">
              <label for="dataset-name">Name:</label>
              <input type="text" id="dataset-name" placeholder="Enter a name for this data set">
            </div>
            <div class="form-group">
              <label for="dataset-description">Description (optional):</label>
              <textarea id="dataset-description" rows="3" placeholder="Enter a description for this data set"></textarea>
            </div>
            <div class="buttons">
              <button id="confirm-save-dataset" class="primary-button">Save</button>
              <button id="cancel-save-dataset" class="secondary-button">Cancel</button>
            </div>
          </div>
        </div>
      `;
      document.body.appendChild(saveModal);
      
      // Add event listeners
      const closeButton = saveModal.querySelector('.close-modal');
      const confirmButton = saveModal.querySelector('#confirm-save-dataset');
      const cancelButton = saveModal.querySelector('#cancel-save-dataset');
      
      closeButton.addEventListener('click', function() {
        saveModal.style.display = 'none';
      });
      
      cancelButton.addEventListener('click', function() {
        saveModal.style.display = 'none';
      });
      
      confirmButton.addEventListener('click', function() {
        const nameInput = document.getElementById('dataset-name');
        const descriptionInput = document.getElementById('dataset-description');
        
        const name = nameInput.value.trim();
        const description = descriptionInput.value.trim();
        
        if (!name) {
          CommonUtils.showStatus('Please enter a name for the data set', 'error');
          return;
        }
        
        // Get current data
        let currentData = '';
        if (tableData.records.length > 0) {
          // Use the parsed data if available
          currentData = convertTableDataToCsv(tableData);
        } else {
          // Otherwise use the raw input
          currentData = bulkRecipientsInput.value.trim();
        }
        
        if (!currentData) {
          CommonUtils.showStatus('No recipient data to save', 'error');
          return;
        }
        
        // Get selected format
        let selectedFormat = 'auto';
        formatRadios.forEach(radio => {
          if (radio.checked) {
            selectedFormat = radio.value;
          }
        });
        
        // Create data set object
        const dataSet = {
          id: Date.now().toString(),
          name: name,
          description: description,
          format: selectedFormat,
          data: currentData,
          recordCount: tableData.records.length,
          dateCreated: new Date().toISOString()
        };
        
        // Save to storage
        chrome.storage.local.get(['recipientDataSets'], function(result) {
          const dataSets = result.recipientDataSets || [];
          dataSets.push(dataSet);
          
          chrome.storage.local.set({ recipientDataSets: dataSets }, function() {
            CommonUtils.showStatus(`Data set "${name}" saved successfully`, 'success');
            saveModal.style.display = 'none';
            nameInput.value = '';
            descriptionInput.value = '';
          });
        });
      });
      
      // Close when clicking outside the modal
      window.addEventListener('click', function(event) {
        if (event.target === saveModal) {
          saveModal.style.display = 'none';
        }
      });
    }
    
    // Show the modal
    saveModal.style.display = 'block';
  }
  
  // Convert table data to CSV format
  function convertTableDataToCsv(data) {
    if (!data.headers || !data.records || data.headers.length === 0) {
      return '';
    }
    
    // Create header row
    const csvRows = [data.headers.join(',')];
    
    // Add data rows
    data.records.forEach(record => {
      const rowValues = data.headers.map(header => {
        // Escape commas and quotes in values
        const value = record[header] || '';
        if (value.includes(',') || value.includes('"')) {
          return `"${value.replace(/"/g, '""')}"`;
        }
        return value;
      });
      csvRows.push(rowValues.join(','));
    });
    
    return csvRows.join('\n');
  }
  
  // Show data sets manager modal
  function showDataSetsManager() {
    // Create modal for managing data sets
    let manageModal = document.getElementById('manage-datasets-modal');
    if (!manageModal) {
      manageModal = document.createElement('div');
      manageModal.id = 'manage-datasets-modal';
      manageModal.className = 'modal';
      manageModal.innerHTML = `
        <div class="modal-content" style="width: 90%; max-width: 800px;">
          <div class="modal-header">
            <span class="close-modal">&times;</span>
            <h3>Manage Recipient Data Sets</h3>
          </div>
          <div class="modal-body">
            <div id="datasets-list-container" style="max-height: 400px; overflow-y: auto;">
              <table id="datasets-table" class="datasets-table" style="width: 100%;">
                <thead>
                  <tr>
                    <th style="width: 30px;"><input type="checkbox" id="select-all-datasets"></th>
                    <th>Name</th>
                    <th>Description</th>
                    <th>Records</th>
                    <th>Date Created</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody id="datasets-list">
                  <!-- Data sets will be listed here -->
                </tbody>
              </table>
            </div>
            <div class="no-datasets" id="no-datasets-message" style="display: none; text-align: center; padding: 20px;">
              No saved data sets found.
            </div>
            <div class="buttons" style="margin-top: 20px;">
              <button id="merge-selected-datasets" class="primary-button">Merge Selected</button>
              <button id="delete-selected-datasets" class="secondary-button">Delete Selected</button>
              <button id="close-datasets-manager" class="secondary-button">Close</button>
            </div>
          </div>
        </div>
      `;
      document.body.appendChild(manageModal);
      
      // Add event listeners
      const closeButton = manageModal.querySelector('.close-modal');
      const closeManagerButton = manageModal.querySelector('#close-datasets-manager');
      const mergeButton = manageModal.querySelector('#merge-selected-datasets');
      const deleteButton = manageModal.querySelector('#delete-selected-datasets');
      const selectAllCheckbox = manageModal.querySelector('#select-all-datasets');
      
      closeButton.addEventListener('click', function() {
        manageModal.style.display = 'none';
      });
      
      closeManagerButton.addEventListener('click', function() {
        manageModal.style.display = 'none';
      });
      
      selectAllCheckbox.addEventListener('change', function() {
        const checkboxes = manageModal.querySelectorAll('.dataset-checkbox');
        checkboxes.forEach(checkbox => {
          checkbox.checked = selectAllCheckbox.checked;
        });
      });
      
      mergeButton.addEventListener('click', function() {
        mergeSelectedDataSets();
      });
      
      deleteButton.addEventListener('click', function() {
        deleteSelectedDataSets();
      });
      
      // Close when clicking outside the modal
      window.addEventListener('click', function(event) {
        if (event.target === manageModal) {
          manageModal.style.display = 'none';
        }
      });
    }
    
    // Load and display data sets
    loadDataSets();
    
    // Show the modal
    manageModal.style.display = 'block';
  }
  
  // Load data sets from storage and display them
  function loadDataSets() {
    const datasetsList = document.getElementById('datasets-list');
    const noDataSetsMessage = document.getElementById('no-datasets-message');
    
    chrome.storage.local.get(['recipientDataSets'], function(result) {
      const dataSets = result.recipientDataSets || [];
      
      if (dataSets.length === 0) {
        datasetsList.innerHTML = '';
        noDataSetsMessage.style.display = 'block';
        return;
      }
      
      noDataSetsMessage.style.display = 'none';
      
      // Sort data sets by date (newest first)
      dataSets.sort((a, b) => new Date(b.dateCreated) - new Date(a.dateCreated));
      
      // Generate HTML for each data set
      datasetsList.innerHTML = dataSets.map((dataSet, index) => {
        const date = new Date(dataSet.dateCreated);
        const formattedDate = `${date.toLocaleDateString()} ${date.toLocaleTimeString()}`;
        
        return `
          <tr data-id="${dataSet.id}">
            <td><input type="checkbox" class="dataset-checkbox" data-id="${dataSet.id}"></td>
            <td>${dataSet.name}</td>
            <td>${dataSet.description || '-'}</td>
            <td>${dataSet.recordCount || 'Unknown'}</td>
            <td>${formattedDate}</td>
            <td>
              <button class="preview-dataset-btn" data-id="${dataSet.id}">Preview</button>
              <button class="load-dataset-btn" data-id="${dataSet.id}">Load</button>
            </td>
          </tr>
        `;
      }).join('');
      
      // Add event listeners to buttons
      const previewButtons = document.querySelectorAll('.preview-dataset-btn');
      const loadButtons = document.querySelectorAll('.load-dataset-btn');
      
      previewButtons.forEach(button => {
        button.addEventListener('click', function() {
          const dataSetId = this.getAttribute('data-id');
          previewDataSet(dataSetId, dataSets);
        });
      });
      
      loadButtons.forEach(button => {
        button.addEventListener('click', function() {
          const dataSetId = this.getAttribute('data-id');
          loadDataSet(dataSetId, dataSets);
        });
      });
    });
  }
  
  // Preview a data set
  function previewDataSet(dataSetId, dataSets) {
    const dataSet = dataSets.find(ds => ds.id === dataSetId);
    if (!dataSet) return;
    
    // Create modal for preview
    let previewModal = document.getElementById('preview-dataset-modal');
    if (!previewModal) {
      previewModal = document.createElement('div');
      previewModal.id = 'preview-dataset-modal';
      previewModal.className = 'modal';
      previewModal.innerHTML = `
        <div class="modal-content">
          <div class="modal-header">
            <span class="close-modal">&times;</span>
            <h3>Data Set Preview</h3>
          </div>
          <div class="modal-body">
            <div class="dataset-info">
              <p><strong>Name:</strong> <span id="preview-dataset-name"></span></p>
              <p><strong>Description:</strong> <span id="preview-dataset-description"></span></p>
              <p><strong>Records:</strong> <span id="preview-dataset-records"></span></p>
              <p><strong>Date Created:</strong> <span id="preview-dataset-date"></span></p>
            </div>
            <div class="form-group">
              <label>Data Preview:</label>
              <textarea id="preview-dataset-data" rows="10" readonly style="font-family: monospace;"></textarea>
            </div>
            <div class="buttons">
              <button id="close-preview-dataset" class="secondary-button">Close</button>
            </div>
          </div>
        </div>
      `;
      document.body.appendChild(previewModal);
      
      // Add event listeners
      const closeButton = previewModal.querySelector('.close-modal');
      const closePreviewButton = previewModal.querySelector('#close-preview-dataset');
      
      closeButton.addEventListener('click', function() {
        previewModal.style.display = 'none';
      });
      
      closePreviewButton.addEventListener('click', function() {
        previewModal.style.display = 'none';
      });
      
      // Close when clicking outside the modal
      window.addEventListener('click', function(event) {
        if (event.target === previewModal) {
          previewModal.style.display = 'none';
        }
      });
    }
    
    // Update preview content
    const nameElement = document.getElementById('preview-dataset-name');
    const descriptionElement = document.getElementById('preview-dataset-description');
    const recordsElement = document.getElementById('preview-dataset-records');
    const dateElement = document.getElementById('preview-dataset-date');
    const dataElement = document.getElementById('preview-dataset-data');
    
    nameElement.textContent = dataSet.name;
    descriptionElement.textContent = dataSet.description || '-';
    recordsElement.textContent = dataSet.recordCount || 'Unknown';
    
    const date = new Date(dataSet.dateCreated);
    dateElement.textContent = `${date.toLocaleDateString()} ${date.toLocaleTimeString()}`;
    
    // Show first 20 lines of data
    const dataLines = dataSet.data.split('\n');
    const previewLines = dataLines.slice(0, 20);
    if (dataLines.length > 20) {
      previewLines.push('... (more data not shown) ...');
    }
    dataElement.value = previewLines.join('\n');
    
    // Show the modal
    previewModal.style.display = 'block';
  }
  
  // Load a data set
  function loadDataSet(dataSetId, dataSets) {
    const dataSet = dataSets.find(ds => ds.id === dataSetId);
    if (!dataSet) return;
    
    if (confirm(`Load data set "${dataSet.name}"? This will replace your current recipient data.`)) {
      // Set the data format
      formatRadios.forEach(radio => {
        radio.checked = radio.value === dataSet.format;
      });
      
      // Set the recipient data
      bulkRecipientsInput.value = dataSet.data;
      
      // Parse and preview the data
      previewBtn.click();
      
      // Close the modal
      const manageModal = document.getElementById('manage-datasets-modal');
      if (manageModal) {
        manageModal.style.display = 'none';
      }
      
      CommonUtils.showStatus(`Data set "${dataSet.name}" loaded successfully`, 'success');
    }
  }
  
  // Merge selected data sets
  function mergeSelectedDataSets() {
    const checkboxes = document.querySelectorAll('.dataset-checkbox:checked');
    if (checkboxes.length === 0) {
      CommonUtils.showStatus('Please select at least one data set to merge', 'error');
      return;
    }
    
    const selectedIds = Array.from(checkboxes).map(checkbox => checkbox.getAttribute('data-id'));
    
    chrome.storage.local.get(['recipientDataSets'], function(result) {
      const dataSets = result.recipientDataSets || [];
      const selectedDataSets = dataSets.filter(ds => selectedIds.includes(ds.id));
      
      if (selectedDataSets.length === 0) return;
      
      // Confirm merge
      if (confirm(`Merge ${selectedDataSets.length} selected data sets with your current data?`)) {
        // Parse each data set
        const allRecords = [];
        let headers = new Set();
        
        // First, collect all unique headers
        selectedDataSets.forEach(dataSet => {
          const parsedData = parseData(dataSet.data, dataSet.format);
          parsedData.headers.forEach(header => headers.add(header));
        });
        
        // Also include current data headers if available
        if (tableData.headers && tableData.headers.length > 0) {
          tableData.headers.forEach(header => headers.add(header));
          // Add current records to the merged set
          allRecords.push(...tableData.records);
        }
        
        headers = Array.from(headers);
        
        // Then, collect all records with the unified headers
        selectedDataSets.forEach(dataSet => {
          const parsedData = parseData(dataSet.data, dataSet.format);
          parsedData.records.forEach(record => {
            allRecords.push(record);
          });
        });
        
        // Create merged table data
        const mergedData = {
          headers: headers,
          records: allRecords
        };
        
        // Update the UI
        tableData = mergedData;
        bulkRecipientsInput.value = convertTableDataToCsv(mergedData);
        
        // Preview the merged data
        previewBtn.click();
        
        // Close the modal
        const manageModal = document.getElementById('manage-datasets-modal');
        if (manageModal) {
          manageModal.style.display = 'none';
        }
        
        CommonUtils.showStatus(`Merged ${selectedDataSets.length} data sets successfully`, 'success');
      }
    });
  }
  
  // Delete selected data sets
  function deleteSelectedDataSets() {
    const checkboxes = document.querySelectorAll('.dataset-checkbox:checked');
    if (checkboxes.length === 0) {
      CommonUtils.showStatus('Please select at least one data set to delete', 'error');
      return;
    }
    
    const selectedIds = Array.from(checkboxes).map(checkbox => checkbox.getAttribute('data-id'));
    
    if (confirm(`Delete ${selectedIds.length} selected data sets? This cannot be undone.`)) {
      chrome.storage.local.get(['recipientDataSets'], function(result) {
        const dataSets = result.recipientDataSets || [];
        const updatedDataSets = dataSets.filter(ds => !selectedIds.includes(ds.id));
        
        chrome.storage.local.set({ recipientDataSets: updatedDataSets }, function() {
          // Reload the data sets list
          loadDataSets();
          
          CommonUtils.showStatus(`Deleted ${selectedIds.length} data sets successfully`, 'success');
        });
      });
    }
  }
  
}

// Export functions
window.BulkEmailTab = {
  init: initBulkEmailTab
};
