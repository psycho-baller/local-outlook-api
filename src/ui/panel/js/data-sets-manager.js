// data-sets-manager.js - Manages recipient data sets functionality

/**
 * DataSetsManager - Handles saving, loading, and managing recipient data sets
 */
class DataSetsManager {
  constructor() {
    this.commonUtils = window.CommonUtils;
    this.currentSort = {
      field: 'dateCreated',
      direction: 'desc'
    };
  }

  /**
   * Initialize event listeners for data set management
   * @param {Object} options - Configuration options
   * @param {HTMLElement} options.saveDataSetBtn - Button to save data set
   * @param {HTMLElement} options.manageDataSetsBtn - Button to manage data sets
   * @param {Function} options.getTableData - Function to get current table data
   * @param {Function} options.getRecipientData - Function to get current recipient data
   * @param {Function} options.getSelectedFormat - Function to get selected format
   * @param {Function} options.parseData - Function to parse data
   * @param {Function} options.previewEmail - Function to preview email to a recipient
   */
  init(options) {
    this.options = options;
    
    // Add event listeners
    if (options.saveDataSetBtn) {
      options.saveDataSetBtn.addEventListener('click', () => this.saveRecipientDataSet());
    }
    
    if (options.manageDataSetsBtn) {
      options.manageDataSetsBtn.addEventListener('click', () => this.showDataSetsManager());
    }
    
    // Initialize modal event listeners when the DOM is fully loaded
    document.addEventListener('DOMContentLoaded', () => {
      this.initializeModalEventListeners();
    });
    
    // Also initialize them now in case DOMContentLoaded already fired
    this.initializeModalEventListeners();
  }

  /**
   * Save current recipient data set with name and description
   */
  saveRecipientDataSet() {
    const saveModal = document.getElementById('save-dataset-modal');
    
    if (saveModal) {
      // Load existing datasets for the dropdown
      this.loadDataSetsForDropdown();
      saveModal.style.display = 'block';
    }
  }
  
  /**
   * Load existing datasets for the dropdown in the save modal
   */
  loadDataSetsForDropdown() {
    const datasetSelect = document.getElementById('dataset-select');
    if (!datasetSelect) return;
    
    // Clear existing options except the first one
    while (datasetSelect.options.length > 1) {
      datasetSelect.remove(1);
    }
    
    // Get datasets from storage
    chrome.storage.local.get(['recipientDataSets'], (result) => {
      const dataSets = result.recipientDataSets || [];
      
      if (dataSets.length === 0) return;
      
      // Sort datasets by name for easier selection
      dataSets.sort((a, b) => a.name.localeCompare(b.name));
      
      // Add options for each dataset
      dataSets.forEach(dataSet => {
        const option = document.createElement('option');
        option.value = dataSet.id;
        option.textContent = dataSet.name;
        option.dataset.description = dataSet.description || '';
        datasetSelect.appendChild(option);
      });
      
      // Add event listener to update name and description when a dataset is selected
      datasetSelect.addEventListener('change', () => {
        const selectedId = datasetSelect.value;
        if (!selectedId) {
          // Clear fields if "Create new dataset" is selected
          document.getElementById('dataset-name').value = '';
          document.getElementById('dataset-description').value = '';
          return;
        }
        
        // Find the selected dataset
        const selectedDataset = dataSets.find(ds => ds.id === selectedId);
        if (selectedDataset) {
          document.getElementById('dataset-name').value = selectedDataset.name;
          document.getElementById('dataset-description').value = selectedDataset.description || '';
        }
      });
    });
  }
  
  /**
   * Confirm saving a data set
   */
  confirmSaveDataSet() {
    const nameInput = document.getElementById('dataset-name');
    const descriptionInput = document.getElementById('dataset-description');
    const datasetSelect = document.getElementById('dataset-select');
    
    const name = nameInput.value.trim();
    const description = descriptionInput.value.trim();
    const selectedDatasetId = datasetSelect ? datasetSelect.value : '';
    
    if (!name) {
      this.commonUtils.showStatus('Please enter a name for the data set', 'error');
      return;
    }
    
    // Get current data
    let currentData = '';
    const tableData = this.options.getTableData();
    
    if (tableData.records.length > 0) {
      // Use the parsed data if available
      currentData = this.convertTableDataToCsv(tableData);
    } else {
      // Otherwise use the raw input
      currentData = this.options.getRecipientData().trim();
    }
    
    if (!currentData) {
      this.commonUtils.showStatus('No recipient data to save', 'error');
      return;
    }
    
    // Get selected format
    const selectedFormat = this.options.getSelectedFormat();
    
    // Save to storage
    chrome.storage.local.get(['recipientDataSets'], (result) => {
      const dataSets = result.recipientDataSets || [];
      
      if (selectedDatasetId) {
        // Update existing dataset
        const existingDatasetIndex = dataSets.findIndex(ds => ds.id === selectedDatasetId);
        
        if (existingDatasetIndex !== -1) {
          // Update the existing dataset
          dataSets[existingDatasetIndex] = {
            ...dataSets[existingDatasetIndex],
            name: name,
            description: description,
            format: selectedFormat,
            data: currentData,
            recordCount: tableData.records.length,
            lastModified: new Date().toISOString()
          };
          
          chrome.storage.local.set({ recipientDataSets: dataSets }, () => {
            this.commonUtils.showStatus(`Data set "${name}" updated successfully`, 'success');
            document.getElementById('save-dataset-modal').style.display = 'none';
            nameInput.value = '';
            descriptionInput.value = '';
            datasetSelect.value = '';
          });
          return;
        }
      }
      
      // Create new dataset
      const newDataSet = {
        id: Date.now().toString(),
        name: name,
        description: description,
        format: selectedFormat,
        data: currentData,
        recordCount: tableData.records.length,
        dateCreated: new Date().toISOString()
      };
      
      dataSets.push(newDataSet);
      
      chrome.storage.local.set({ recipientDataSets: dataSets }, () => {
        this.commonUtils.showStatus(`Data set "${name}" saved successfully`, 'success');
        document.getElementById('save-dataset-modal').style.display = 'none';
        nameInput.value = '';
        descriptionInput.value = '';
        if (datasetSelect) datasetSelect.value = '';
      });
    });
  }

  /**
   * Convert table data to CSV format
   * @param {Object} data - Table data with headers and records
   * @returns {string} CSV formatted data
   */
  convertTableDataToCsv(data) {
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

  /**
   * Show data sets manager modal
   */
  showDataSetsManager() {
    // Load and display data sets
    this.loadDataSets();
    
    // Show the modal
    const manageModal = document.getElementById('manage-datasets-modal');
    if (manageModal) {
      manageModal.style.display = 'block';
    }
  }
  

  
  /**
   * Initialize event listeners for modals
   */
  initializeModalEventListeners() {
    // Setup event listeners for the data sets manager modal
    const manageModal = document.getElementById('manage-datasets-modal');
    if (!manageModal) return;
    
    // The modal HTML is now in the HTML file, so we just need to add event listeners
    const closeButton = manageModal.querySelector('.close-modal');
    const closeManagerButton = document.getElementById('close-datasets-manager');
    const mergeButton = document.getElementById('merge-selected-datasets');
    const intersectButton = document.getElementById('intersect-selected-datasets');
    const subtractButton = document.getElementById('subtract-selected-datasets');
    const deleteButton = document.getElementById('delete-selected-datasets');
    const selectAllCheckbox = document.getElementById('select-all-datasets');
    const sortableHeaders = manageModal.querySelectorAll('.sortable');
    
    if (closeButton) {
      closeButton.addEventListener('click', () => {
        manageModal.style.display = 'none';
      });
    }
    
    if (closeManagerButton) {
      closeManagerButton.addEventListener('click', () => {
        manageModal.style.display = 'none';
      });
    }
    
    if (selectAllCheckbox) {
      selectAllCheckbox.addEventListener('change', () => {
        const checkboxes = manageModal.querySelectorAll('.dataset-checkbox');
        checkboxes.forEach(checkbox => {
          checkbox.checked = selectAllCheckbox.checked;
        });
      });
    }
    
    if (mergeButton) {
      mergeButton.addEventListener('click', () => {
        this.mergeSelectedDataSets();
      });
    }
    
    if (intersectButton) {
      intersectButton.addEventListener('click', () => {
        this.intersectSelectedDataSets();
      });
    }
    
    if (subtractButton) {
      subtractButton.addEventListener('click', () => {
        this.subtractSelectedDataSets();
      });
    }
    
    if (deleteButton) {
      deleteButton.addEventListener('click', () => {
        this.deleteSelectedDataSets();
      });
    }
    
    // Add event listeners for sortable headers
    sortableHeaders.forEach(header => {
      header.addEventListener('click', (e) => {
        e.preventDefault();
        const sortField = header.getAttribute('data-sort');
        this.sortDataSets(sortField);
      });
      
      // Set default sort if specified
      const defaultSort = header.getAttribute('data-sort-default');
      if (defaultSort && header.getAttribute('data-sort') === this.currentSort.field) {
        header.classList.add(`sort-${this.currentSort.direction}`);
      }
    });
    
    // Close when clicking outside the modal
    manageModal.addEventListener('click', (event) => {
      if (event.target === manageModal) {
        manageModal.style.display = 'none';
      }
    });
    
    // Setup event listeners for save dataset modal
    const saveModal = document.getElementById('save-dataset-modal');
    if (saveModal) {
      const saveCloseButton = saveModal.querySelector('.close-modal');
      const confirmButton = document.getElementById('confirm-save-dataset');
      const cancelButton = document.getElementById('cancel-save-dataset');
      
      if (saveCloseButton) {
        saveCloseButton.addEventListener('click', () => {
          saveModal.style.display = 'none';
        });
      }
      
      if (cancelButton) {
        cancelButton.addEventListener('click', () => {
          saveModal.style.display = 'none';
        });
      }
      
      if (confirmButton) {
        confirmButton.addEventListener('click', () => {
          this.confirmSaveDataSet();
        });
      }
      
      saveModal.addEventListener('click', (event) => {
        if (event.target === saveModal) {
          saveModal.style.display = 'none';
        }
      });
    }
    
    // Setup event listeners for preview dataset modal
    const previewModal = document.getElementById('preview-dataset-modal');
    if (previewModal) {
      const previewCloseButton = previewModal.querySelector('.close-modal');
      const closePreviewButton = document.getElementById('close-preview-dataset');
      
      if (previewCloseButton) {
        previewCloseButton.addEventListener('click', () => {
          previewModal.style.display = 'none';
        });
      }
      
      if (closePreviewButton) {
        closePreviewButton.addEventListener('click', () => {
          previewModal.style.display = 'none';
        });
      }
      
      previewModal.addEventListener('click', (event) => {
        if (event.target === previewModal) {
          previewModal.style.display = 'none';
        }
      });
    }
  }

  /**
   * Sort data sets by the specified field
   * @param {string} field - Field to sort by
   */
  sortDataSets(field) {
    // If clicking the same header, toggle direction
    if (field === this.currentSort.field) {
      this.currentSort.direction = this.currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
      this.currentSort.field = field;
      this.currentSort.direction = 'asc';
    }
    
    // Update UI to show sort direction
    const sortableHeaders = document.querySelectorAll('.sortable');
    sortableHeaders.forEach(header => {
      header.classList.remove('sort-asc', 'sort-desc');
      if (header.getAttribute('data-sort') === field) {
        header.classList.add(`sort-${this.currentSort.direction}`);
      }
    });
    
    // Reload data sets with new sort
    this.loadDataSets();
  }

  /**
   * Load data sets from storage and display them
   */
  loadDataSets() {
    const datasetsList = document.getElementById('datasets-list');
    const noDataSetsMessage = document.getElementById('no-datasets-message');
    
    chrome.storage.local.get(['recipientDataSets'], (result) => {
      const dataSets = result.recipientDataSets || [];
      
      if (dataSets.length === 0) {
        datasetsList.innerHTML = '';
        noDataSetsMessage.style.display = 'block';
        return;
      }
      
      noDataSetsMessage.style.display = 'none';
      
      // Sort data sets based on current sort settings
      dataSets.sort((a, b) => {
        let aValue = a[this.currentSort.field];
        let bValue = b[this.currentSort.field];
        
        // Handle special cases
        if (this.currentSort.field === 'dateCreated') {
          aValue = new Date(aValue);
          bValue = new Date(bValue);
        } else if (this.currentSort.field === 'recordCount') {
          aValue = Number(aValue || 0);
          bValue = Number(bValue || 0);
        } else {
          // For string fields like name and description
          aValue = String(aValue || '').toLowerCase();
          bValue = String(bValue || '').toLowerCase();
        }
        
        // Compare based on direction
        if (this.currentSort.direction === 'asc') {
          return aValue > bValue ? 1 : aValue < bValue ? -1 : 0;
        } else {
          return aValue < bValue ? 1 : aValue > bValue ? -1 : 0;
        }
      });
      
      // Generate HTML for each data set
      datasetsList.innerHTML = dataSets.map((dataSet) => {
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
        button.addEventListener('click', () => {
          const dataSetId = button.getAttribute('data-id');
          this.previewDataSet(dataSetId, dataSets);
        });
      });
      
      loadButtons.forEach(button => {
        button.addEventListener('click', () => {
          const dataSetId = button.getAttribute('data-id');
          this.loadDataSet(dataSetId, dataSets);
        });
      });
    });
  }

  /**
   * Preview a data set
   * @param {string} dataSetId - ID of the data set to preview
   * @param {Array} dataSets - Array of data sets
   */
  previewDataSet(dataSetId, dataSets) {
    const dataSet = dataSets.find(ds => ds.id === dataSetId);
    if (!dataSet) return;
    
    const previewModal = document.getElementById('preview-dataset-modal');
    
    if (previewModal) {
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
      return;
    }
    
    // The modal HTML is now in the HTML file, so we just need to add event listeners
    const closeButton = document.querySelector('#preview-dataset-modal .close-modal');
    const closePreviewButton = document.querySelector('#close-preview-dataset');
    
    closeButton.addEventListener('click', () => {
      document.getElementById('preview-dataset-modal').style.display = 'none';
    });
    
    closePreviewButton.addEventListener('click', () => {
      document.getElementById('preview-dataset-modal').style.display = 'none';
    });
    
    // Close when clicking outside the modal
    window.addEventListener('click', (event) => {
      const modal = document.getElementById('preview-dataset-modal');
      if (event.target === modal) {
        modal.style.display = 'none';
      }
    });
    
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
    document.getElementById('preview-dataset-modal').style.display = 'block';
  }

  /**
   * Load a data set
   * @param {string} dataSetId - ID of the data set to load
   * @param {Array} dataSets - Array of data sets
   */
  loadDataSet(dataSetId, dataSets) {
    const dataSet = dataSets.find(ds => ds.id === dataSetId);
    if (!dataSet) return;
    
    if (confirm(`Load data set "${dataSet.name}"? This will replace your current recipient data.`)) {
      // Set the data format
      const formatRadios = document.querySelectorAll('input[name="data-format"]');
      formatRadios.forEach(radio => {
        radio.checked = radio.value === dataSet.format;
      });
      
      // Set the recipient data
      const bulkRecipientsInput = document.getElementById('bulk-recipients');
      bulkRecipientsInput.value = dataSet.data;
      
      // Parse and preview the data
      const previewBtn = document.getElementById('previewBtn');
      previewBtn.click();
      
      // Close the modal
      const manageModal = document.getElementById('manage-datasets-modal');
      if (manageModal) {
        manageModal.style.display = 'none';
      }
      
      this.commonUtils.showStatus(`Data set "${dataSet.name}" loaded successfully`, 'success');
    }
  }

  /**
   * Find the email field in headers
   * @param {Array} headers - Array of header names
   * @returns {string|null} - The email field name or null if not found
   * @private
   */
  _findEmailField(headers) {
    return headers.find(header => 
      ['email', 'emailaddress', 'e-mail', 'mail'].includes(header.toLowerCase())
    ) || null;
  }

  /**
   * Process data sets and prepare for operation
   * @param {string} operation - The operation type ('union', 'intersection', 'subtraction')
   * @private
   */
  _processSelectedDataSets(operation) {
    const checkboxes = document.querySelectorAll('.dataset-checkbox:checked');
    if (checkboxes.length === 0) {
      this.commonUtils.showStatus(`Please select at least one data set for ${operation}`, 'error');
      return null;
    }
    
    // For subtraction, we need exactly two data sets
    if (operation === 'subtraction' && checkboxes.length !== 2) {
      this.commonUtils.showStatus('Please select exactly two data sets for subtraction', 'error');
      return null;
    }
    
    // For intersection, we need at least two data sets
    if (operation === 'intersection' && checkboxes.length < 2) {
      this.commonUtils.showStatus('Please select at least two data sets for intersection', 'error');
      return null;
    }
    
    const selectedIds = Array.from(checkboxes).map(checkbox => checkbox.getAttribute('data-id'));
    
    return new Promise((resolve) => {
      chrome.storage.local.get(['recipientDataSets'], (result) => {
        const dataSets = result.recipientDataSets || [];
        const selectedDataSets = dataSets.filter(ds => selectedIds.includes(ds.id));
        
        if (selectedDataSets.length === 0) {
          resolve(null);
          return;
        }
        
        // Get current table data
        const tableData = this.options.getTableData();
        
        // Parse all data sets
        const parsedDataSets = selectedDataSets.map(dataSet => {
          return {
            id: dataSet.id,
            name: dataSet.name,
            format: dataSet.format,
            parsedData: this.options.parseData(dataSet.data, dataSet.format)
          };
        });
        
        // Collect all unique headers
        let headers = new Set();
        parsedDataSets.forEach(ds => {
          ds.parsedData.headers.forEach(header => headers.add(header));
        });
        
        // Also include current data headers if available
        if (tableData.headers && tableData.headers.length > 0) {
          tableData.headers.forEach(header => headers.add(header));
        }
        
        headers = Array.from(headers);
        
        // Find email field in headers
        const emailField = this._findEmailField(headers);
        if (!emailField && (operation === 'intersection' || operation === 'subtraction')) {
          this.commonUtils.showStatus(`No email field found. ${operation} operation requires an email field.`, 'error');
          resolve(null);
          return;
        }
        
        resolve({
          selectedDataSets,
          parsedDataSets,
          headers,
          emailField,
          tableData
        });
      });
    });
  }

  /**
   * Merge selected data sets (Union operation)
   */
  async mergeSelectedDataSets() {
    const data = await this._processSelectedDataSets('union');
    if (!data) return;
    
    const { selectedDataSets, parsedDataSets, headers, tableData } = data;
    
    // Confirm operation
    if (confirm(`Merge (union) ${selectedDataSets.length} selected data sets with your current data?`)) {
      // Collect all records
      const allRecords = [];
      
      // Add current records if available
      if (tableData.headers && tableData.headers.length > 0) {
        allRecords.push(...tableData.records);
      }
      
      // Add all records from selected data sets
      parsedDataSets.forEach(ds => {
        allRecords.push(...ds.parsedData.records);
      });
      
      // Create merged table data
      const mergedData = {
        headers: headers,
        records: allRecords
      };
      
      // Update the UI
      const bulkRecipientsInput = document.getElementById('bulk-recipients');
      bulkRecipientsInput.value = this.convertTableDataToCsv(mergedData);
      
      // Preview the merged data
      const previewBtn = document.getElementById('previewBtn');
      previewBtn.click();
      
      // Close the modal
      const manageModal = document.getElementById('manage-datasets-modal');
      if (manageModal) {
        manageModal.style.display = 'none';
      }
      
      this.commonUtils.showStatus(`Union of ${selectedDataSets.length} data sets completed successfully`, 'success');
    }
  }
  
  /**
   * Intersect selected data sets (keep only records that exist in all data sets, matching by email)
   */
  async intersectSelectedDataSets() {
    const data = await this._processSelectedDataSets('intersection');
    if (!data) return;
    
    const { selectedDataSets, parsedDataSets, headers, emailField, tableData } = data;
    
    // Confirm operation
    if (confirm(`Find intersection of ${selectedDataSets.length} selected data sets? This will keep only records with matching emails across all data sets.`)) {
      // Create email-to-record maps for each data set
      const dataSetsEmailMaps = parsedDataSets.map(ds => {
        const emailMap = new Map();
        ds.parsedData.records.forEach(record => {
          const email = record[emailField]?.toLowerCase();
          if (email) {
            emailMap.set(email, record);
          }
        });
        return emailMap;
      });
      
      // Also include current data if available
      let currentDataEmailMap = new Map();
      if (tableData.headers && tableData.headers.length > 0) {
        tableData.records.forEach(record => {
          const email = record[emailField]?.toLowerCase();
          if (email) {
            currentDataEmailMap.set(email, record);
          }
        });
        // Add current data email map to the list
        dataSetsEmailMaps.push(currentDataEmailMap);
      }
      
      // Find emails that exist in all data sets
      let commonEmails = null;
      
      dataSetsEmailMaps.forEach(emailMap => {
        const emails = Array.from(emailMap.keys());
        if (commonEmails === null) {
          commonEmails = new Set(emails);
        } else {
          commonEmails = new Set(emails.filter(email => commonEmails.has(email)));
        }
      });
      
      // Create records for intersection
      const intersectionRecords = [];
      
      if (commonEmails && commonEmails.size > 0) {
        // Use the first data set as the base for record data
        const baseEmailMap = dataSetsEmailMaps[0];
        
        // Create a merged record for each common email
        commonEmails.forEach(email => {
          const mergedRecord = {};
          
          // Start with the record from the first data set
          const baseRecord = baseEmailMap.get(email);
          Object.assign(mergedRecord, baseRecord);
          
          // Merge in data from other data sets (if fields don't already exist)
          for (let i = 1; i < dataSetsEmailMaps.length; i++) {
            const otherRecord = dataSetsEmailMaps[i].get(email);
            if (otherRecord) {
              Object.keys(otherRecord).forEach(key => {
                if (mergedRecord[key] === undefined || mergedRecord[key] === '') {
                  mergedRecord[key] = otherRecord[key];
                }
              });
            }
          }
          
          intersectionRecords.push(mergedRecord);
        });
      }
      
      // Create intersection table data
      const intersectionData = {
        headers: headers,
        records: intersectionRecords
      };
      
      // Update the UI
      const bulkRecipientsInput = document.getElementById('bulk-recipients');
      bulkRecipientsInput.value = this.convertTableDataToCsv(intersectionData);
      
      // Preview the intersection data
      const previewBtn = document.getElementById('previewBtn');
      previewBtn.click();
      
      // Close the modal
      const manageModal = document.getElementById('manage-datasets-modal');
      if (manageModal) {
        manageModal.style.display = 'none';
      }
      
      this.commonUtils.showStatus(`Intersection found ${intersectionRecords.length} common records across ${selectedDataSets.length} data sets`, 'success');
    }
  }
  
  /**
   * Subtract second data set from first data set (remove records from first that exist in second, matching by email)
   */
  async subtractSelectedDataSets() {
    const data = await this._processSelectedDataSets('subtraction');
    if (!data) return;
    
    const { selectedDataSets, parsedDataSets, headers, emailField, tableData } = data;
    
    // For subtraction, we need exactly two data sets
    if (parsedDataSets.length !== 2) {
      this.commonUtils.showStatus('Please select exactly two data sets for subtraction', 'error');
      return;
    }
    
    // Confirm operation
    if (confirm(`Subtract records from "${selectedDataSets[1].name}" from "${selectedDataSets[0].name}"? This will remove records with matching emails.`)) {
      // Create email-to-record maps for both data sets
      const firstDataSet = parsedDataSets[0];
      const secondDataSet = parsedDataSets[1];
      
      // Get emails from second data set
      const secondDataSetEmails = new Set();
      secondDataSet.parsedData.records.forEach(record => {
        const email = record[emailField]?.toLowerCase();
        if (email) {
          secondDataSetEmails.add(email);
        }
      });
      
      // Filter records from first data set that don't exist in second data set
      const resultRecords = firstDataSet.parsedData.records.filter(record => {
        const email = record[emailField]?.toLowerCase();
        return !email || !secondDataSetEmails.has(email);
      });
      
      // Create subtraction table data
      const subtractionData = {
        headers: headers,
        records: resultRecords
      };
      
      // Update the UI
      const bulkRecipientsInput = document.getElementById('bulk-recipients');
      bulkRecipientsInput.value = this.convertTableDataToCsv(subtractionData);
      
      // Preview the subtraction data
      const previewBtn = document.getElementById('previewBtn');
      previewBtn.click();
      
      // Close the modal
      const manageModal = document.getElementById('manage-datasets-modal');
      if (manageModal) {
        manageModal.style.display = 'none';
      }
      
      const removedCount = firstDataSet.parsedData.records.length - resultRecords.length;
      this.commonUtils.showStatus(`Subtraction completed: ${removedCount} records removed, ${resultRecords.length} records remaining`, 'success');
    }
  }

  /**
   * Delete selected data sets
   */
  deleteSelectedDataSets() {
    const checkboxes = document.querySelectorAll('.dataset-checkbox:checked');
    if (checkboxes.length === 0) {
      this.commonUtils.showStatus('Please select at least one data set to delete', 'error');
      return;
    }
    
    const selectedIds = Array.from(checkboxes).map(checkbox => checkbox.getAttribute('data-id'));
    
    if (confirm(`Delete ${selectedIds.length} selected data sets? This cannot be undone.`)) {
      chrome.storage.local.get(['recipientDataSets'], (result) => {
        const dataSets = result.recipientDataSets || [];
        const updatedDataSets = dataSets.filter(ds => !selectedIds.includes(ds.id));
        
        chrome.storage.local.set({ recipientDataSets: updatedDataSets }, () => {
          // Reload the data sets list
          this.loadDataSets();
          
          this.commonUtils.showStatus(`Deleted ${selectedIds.length} data sets successfully`, 'success');
        });
      });
    }
  }
}

// Export the class
window.DataSetsManager = new DataSetsManager();
