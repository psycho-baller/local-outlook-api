// bulk-email-tab.js - Bulk email tab functionality

// Initialize the bulk email tab
function initBulkEmailTab() {
  // Initialize the bulk email tab
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
  const addRowBtn = document.getElementById('addRowBtn');
  const removeSelectedBtn = document.getElementById('removeSelectedBtn');
  const applyChangesBtn = document.getElementById('applyChangesBtn');
  const saveDataSetBtn = document.getElementById('saveDataSetBtn');
  const manageDataSetsBtn = document.getElementById('manageDataSetsBtn');
  const saveTemplateBtn = document.getElementById('saveTemplateBtn');
  const loadTemplateBtn = document.getElementById('loadTemplateBtn');
  
  // Store the parsed data for editing
  let tableData = {
    headers: [],
    records: []
  };
  
  // Store the original records for filtering
  let originalRecords = [];
  
  // Store current filters
  let currentFilters = {};
  
  // Store checkbox states for each record
  let checkboxStates = {};
  
  // Store active filter element to restore focus
  let activeFilterElement = null;
  let activeFilterHeader = null;
  let activeFilterValue = '';
  
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
    
    // Store original records for filtering
    originalRecords = [...parsedData.records];
    
    // Initialize checkbox states for all records (default to checked)
    checkboxStates = {};
    originalRecords.forEach((_, index) => {
      checkboxStates[index] = true;
    });
    
    // Reset filters
    currentFilters = {};
    
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
  
  // Simple debounce function
  function debounce(func, wait) {
    let timeout;
    return function(...args) {
      const context = this;
      clearTimeout(timeout);
      timeout = setTimeout(() => func.apply(context, args), wait);
    };
  }
  
  // Apply filters to the records - debounced version
  const debouncedApplyFilters = debounce(function() {
    // Start with all original records
    let filteredRecords = [...originalRecords];
    
    // Save current checkbox states before filtering
    saveCheckboxStates();
    
    // Apply each filter
    Object.keys(currentFilters).forEach(header => {
      const filterValue = currentFilters[header].toLowerCase();
      if (filterValue) {
        filteredRecords = filteredRecords.filter(record => {
          const cellValue = String(record[header] || '').toLowerCase();
          return cellValue.includes(filterValue);
        });
      }
    });
    
    // Update tableData with filtered records
    tableData.records = filteredRecords;
    
    // Re-render the table while preserving filter focus
    renderEditableTable(true);
  }, 250); // 250ms debounce time
  
  // Apply filters to the records
  function applyFilters() {
    debouncedApplyFilters();
  }

  // Render the editable table with the current data
  function renderEditableTable(skipFilterUpdate = false) {
    // Clear the table
    previewTable.innerHTML = '';
    
    // Create header row
    const headerRow = document.createElement('tr');
    
    // Add checkbox column header with Select/Deselect All checkbox
    const checkboxHeader = document.createElement('th');
    checkboxHeader.style.width = '30px';
    
    // Create the Select/Deselect All checkbox
    const selectAllCheckbox = document.createElement('input');
    selectAllCheckbox.type = 'checkbox';
    selectAllCheckbox.id = 'select-all-checkbox';
    selectAllCheckbox.checked = true;
    selectAllCheckbox.title = 'Select/Deselect All';
    
    // Add event listener to the checkbox
    selectAllCheckbox.addEventListener('change', function() {
      const checkboxes = document.querySelectorAll('.row-checkbox');
      checkboxes.forEach(checkbox => {
        checkbox.checked = selectAllCheckbox.checked;
        
        // Update stored checkbox states
        const originalIndex = parseInt(checkbox.dataset.originalIndex);
        if (originalIndex !== -1) {
          checkboxStates[originalIndex] = selectAllCheckbox.checked;
        }
      });
    });
    
    checkboxHeader.appendChild(selectAllCheckbox);
    headerRow.appendChild(checkboxHeader);
    
    // Add editable headers with sort functionality
    tableData.headers.forEach((header, index) => {
      const th = document.createElement('th');
      
      // Create header container for name, sort indicator, and filter
      const headerContainer = document.createElement('div');
      headerContainer.className = 'header-container';
      headerContainer.style.display = 'flex';
      headerContainer.style.flexDirection = 'column';
      
      // Create top row with header name and sort indicator
      const headerTopRow = document.createElement('div');
      headerTopRow.style.display = 'flex';
      headerTopRow.style.alignItems = 'center';
      headerTopRow.style.marginBottom = '5px';
      
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
        const newHeader = this.value.trim();
        
        if (newHeader && newHeader !== oldHeader) {
          // Update the header in the tableData
          tableData.headers[index] = newHeader;
          
          // Update the property name in all records
          tableData.records.forEach(record => {
            record[newHeader] = record[oldHeader];
            delete record[oldHeader];
          });
          
          // Update original records too
          originalRecords.forEach(record => {
            record[newHeader] = record[oldHeader];
            delete record[oldHeader];
          });
          
          // Update filter key if it exists
          if (currentFilters[oldHeader]) {
            currentFilters[newHeader] = currentFilters[oldHeader];
            delete currentFilters[oldHeader];
          }
          
          CommonUtils.showStatus(`Column renamed from "${oldHeader}" to "${newHeader}"`, 'success');
        } else if (!newHeader) {
          // Reset to old header if empty
          this.value = oldHeader;
          CommonUtils.showStatus('Header name cannot be empty', 'error');
        }
      });
      
      // Create sort indicator
      const sortIndicator = document.createElement('span');
      sortIndicator.className = 'sort-indicator';
      sortIndicator.innerHTML = '&#8597;'; // Up/down arrow
      sortIndicator.style.cursor = 'pointer';
      sortIndicator.style.marginLeft = '5px';
      sortIndicator.title = 'Sort by this column';
      
      // Add click event for sorting
      sortIndicator.addEventListener('click', function() {
        sortTableData(index);
      });
      
      // If this is the current sort column, update the indicator
      if (currentSort.column === index) {
        sortIndicator.innerHTML = currentSort.direction === 'asc' ? '&#8593;' : '&#8595;';
      }
      
      headerTopRow.appendChild(input);
      headerTopRow.appendChild(sortIndicator);
      
      // Create filter input
      const filterInput = document.createElement('input');
      filterInput.type = 'text';
      filterInput.placeholder = 'Filter...';
      filterInput.className = 'filter-input';
      filterInput.dataset.header = header;
      filterInput.style.width = '100%';
      filterInput.style.boxSizing = 'border-box';
      filterInput.style.padding = '2px 5px';
      filterInput.style.fontSize = '12px';
      
      // Set value from current filters or active filter
      if (activeFilterHeader === header && activeFilterValue) {
        filterInput.value = activeFilterValue;
      } else if (!skipFilterUpdate && currentFilters[header]) {
        filterInput.value = currentFilters[header];
      }
      
      // Add event listener for focus to track active filter
      filterInput.addEventListener('focus', function() {
        activeFilterElement = this;
        activeFilterHeader = header;
      });
      
      // Add event listener for filtering
      filterInput.addEventListener('input', function(e) {
        const filterValue = this.value.trim();
        activeFilterValue = this.value; // Store the current value
        
        if (filterValue) {
          currentFilters[header] = filterValue;
        } else {
          delete currentFilters[header];
        }
        
        // Store the selection/cursor position
        const selectionStart = this.selectionStart;
        const selectionEnd = this.selectionEnd;
        
        applyFilters();
        
        // Focus will be restored after render via the activeFilterElement reference
      });
      
      headerContainer.appendChild(headerTopRow);
      headerContainer.appendChild(filterInput);
      th.appendChild(headerContainer);
      headerRow.appendChild(th);
    });
    
    // Add preview column header
    const previewHeader = document.createElement('th');
    previewHeader.textContent = 'Preview';
    previewHeader.style.width = '80px';
    headerRow.appendChild(previewHeader);
    
    // Add header row to table
    previewTable.appendChild(headerRow);
    
    // Update recipient count
    recipientCountSpan.textContent = tableData.records.length;
    
    // Restore focus to active filter input if there was one
    if (activeFilterHeader) {
      // Find the filter input for the active header
      setTimeout(() => {
        const filterInputs = document.querySelectorAll('.filter-input');
        filterInputs.forEach(input => {
          if (input.dataset.header === activeFilterHeader) {
            input.focus();
            // Try to restore cursor position if possible
            try {
              input.setSelectionRange(
                activeFilterValue.length,
                activeFilterValue.length
              );
            } catch (e) {
              // Ignore any errors with selection range
            }
          }
        });
      }, 0);  // Use setTimeout to ensure DOM is fully updated
    }
    
    // Create data rows
    tableData.records.forEach((record, rowIndex) => {
      const row = document.createElement('tr');
      row.dataset.index = rowIndex;
      
      // Add checkbox cell
      const checkboxCell = document.createElement('td');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      
      // Get the original index of this record in the originalRecords array
      const originalIndex = originalRecords.findIndex(r => r === record);
      
      // Set checkbox state from stored states or default to true
      checkbox.checked = originalIndex !== -1 ? 
        (checkboxStates[originalIndex] !== undefined ? checkboxStates[originalIndex] : true) : 
        true;
      
      checkbox.className = 'row-checkbox';
      checkbox.dataset.index = rowIndex;
      checkbox.dataset.originalIndex = originalIndex;
      
      // Add event listener to update checkbox state when changed
      checkbox.addEventListener('change', function() {
        const origIndex = parseInt(this.dataset.originalIndex);
        if (origIndex !== -1) {
          checkboxStates[origIndex] = this.checked;
        }
      });
      
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
  
  // Function to save current checkbox states
  function saveCheckboxStates() {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    checkboxes.forEach(checkbox => {
      const originalIndex = parseInt(checkbox.dataset.originalIndex);
      if (originalIndex !== -1) {
        checkboxStates[originalIndex] = checkbox.checked;
      }
    });
  }
  
  // Note: Select/deselect all functionality is now handled by the checkbox in the table header
  
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
    prepareBulkMail( selectedRecords, emailField, subject, body);
  });

  function prepareBulkMail(selectedRecords, emailField, subject, body) {
    chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
      const currentTab = tabs[0];
      const isOutlook = currentTab && currentTab.url &&
        (currentTab.url.includes('outlook.office.com') ||
          currentTab.url.includes('outlook.live.com'));
  
      if (isOutlook) {
        // Current tab is Outlook, send directly
        processBulkEmails(currentTab.id, selectedRecords, emailField, subject, body);
      } else {
        // Current tab is not Outlook, try to find or create an Outlook tab
        chrome.runtime.sendMessage({ action: 'activateOutlookTab' }, function (result) {
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
  }
  
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
  
  // Initialize the Data Sets Manager
  window.DataSetsManager.init({
    saveDataSetBtn: saveDataSetBtn,
    manageDataSetsBtn: manageDataSetsBtn,
    getTableData: function() { return tableData; },
    getRecipientData: function() { return bulkRecipientsInput.value; },
    getSelectedFormat: function() {
      let selectedFormat = 'auto';
      formatRadios.forEach(radio => {
        if (radio.checked) {
          selectedFormat = radio.value;
        }
      });
      return selectedFormat;
    },
    parseData: parseData,
    previewEmail: showEmailPreview
  });
  
  // The Data Sets Manager functionality has been moved to data-sets-manager.js
  
  // Initialize the Email Templates Manager
  window.EmailTemplatesManager.init({
    saveTemplateBtn: saveTemplateBtn,
    loadTemplateBtn: loadTemplateBtn,
    getSubject: function() { return bulkSubjectInput.value; },
    getBody: function() { return bulkBodyInput.value || bulkBodyEditor.innerHTML; },
    setSubject: function(subject) { bulkSubjectInput.value = subject; },
    setBody: function(body) {
      // Set the body content in both the hidden textarea and the editor
      bulkBodyInput.value = body;
      bulkBodyEditor.innerHTML = body;
      // Save the updated content
      saveBulkData();
    }
  });
}

// Export functions
window.BulkEmailTab = {
  init: initBulkEmailTab
};


