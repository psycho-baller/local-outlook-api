// email-templates-manager.js - Manages email templates functionality

/**
 * EmailTemplatesManager - Handles saving, loading, and managing email templates
 */
class EmailTemplatesManager {
  constructor() {
    this.commonUtils = window.CommonUtils;
    this.currentSort = {
      field: 'dateCreated',
      direction: 'desc'
    };
  }

  /**
   * Initialize event listeners for email template management
   * @param {Object} options - Configuration options
   * @param {HTMLElement} options.saveTemplateBtn - Button to save template
   * @param {HTMLElement} options.loadTemplateBtn - Button to load template
   * @param {Function} options.getSubject - Function to get current subject
   * @param {Function} options.getBody - Function to get current body
   * @param {Function} options.setSubject - Function to set subject
   * @param {Function} options.setBody - Function to set body
   */
  init(options) {
    this.options = options;
    
    // Add event listeners
    if (options.saveTemplateBtn) {
      options.saveTemplateBtn.addEventListener('click', () => this.saveEmailTemplate());
    }
    
    if (options.loadTemplateBtn) {
      options.loadTemplateBtn.addEventListener('click', () => this.showTemplatesManager());
    }
    
    // Initialize modal event listeners when the DOM is fully loaded
    document.addEventListener('DOMContentLoaded', () => {
      this.initializeModalEventListeners();
    });
    
    // Also initialize them now in case DOMContentLoaded already fired
    this.initializeModalEventListeners();
  }

  /**
   * Initialize event listeners for modals
   */
  initializeModalEventListeners() {
    // Save template modal
    const saveTemplateModal = document.getElementById('save-template-modal');
    const confirmSaveTemplateBtn = document.getElementById('confirm-save-template');
    const cancelSaveTemplateBtn = document.getElementById('cancel-save-template');
    
    if (confirmSaveTemplateBtn) {
      confirmSaveTemplateBtn.addEventListener('click', () => this.confirmSaveTemplate());
    }
    
    if (cancelSaveTemplateBtn) {
      cancelSaveTemplateBtn.addEventListener('click', () => {
        if (saveTemplateModal) saveTemplateModal.style.display = 'none';
      });
    }
    
    // Templates manager modal
    const templatesModal = document.getElementById('email-templates-modal');
    const closeTemplatesManagerBtn = document.getElementById('close-templates-manager');
    const deleteSelectedTemplatesBtn = document.getElementById('delete-selected-templates');
    const selectAllTemplatesCheckbox = document.getElementById('select-all-templates');
    
    if (closeTemplatesManagerBtn) {
      closeTemplatesManagerBtn.addEventListener('click', () => {
        if (templatesModal) templatesModal.style.display = 'none';
      });
    }
    
    if (deleteSelectedTemplatesBtn) {
      deleteSelectedTemplatesBtn.addEventListener('click', () => this.deleteSelectedTemplates());
    }
    
    if (selectAllTemplatesCheckbox) {
      selectAllTemplatesCheckbox.addEventListener('change', () => {
        const checkboxes = document.querySelectorAll('#templates-list input[type="checkbox"]');
        checkboxes.forEach(checkbox => {
          checkbox.checked = selectAllTemplatesCheckbox.checked;
        });
      });
    }
    
    // Preview template modal
    const previewTemplateModal = document.getElementById('preview-template-modal');
    const closePreviewTemplateBtn = document.getElementById('close-preview-template');
    const applyTemplateBtn = document.getElementById('apply-template');
    
    if (closePreviewTemplateBtn) {
      closePreviewTemplateBtn.addEventListener('click', () => {
        if (previewTemplateModal) previewTemplateModal.style.display = 'none';
      });
    }
    
    if (applyTemplateBtn) {
      applyTemplateBtn.addEventListener('click', () => this.applyCurrentTemplate());
    }
    
    // Add event listeners to close modals when clicking outside
    window.addEventListener('click', (event) => {
      if (event.target === saveTemplateModal) {
        saveTemplateModal.style.display = 'none';
      } else if (event.target === templatesModal) {
        templatesModal.style.display = 'none';
      } else if (event.target === previewTemplateModal) {
        previewTemplateModal.style.display = 'none';
      }
    });
    
    // Add event listeners to close modals when clicking X
    const closeButtons = document.querySelectorAll('.close-modal');
    closeButtons.forEach(button => {
      button.addEventListener('click', () => {
        const modal = button.closest('.modal');
        if (modal) modal.style.display = 'none';
      });
    });
    
    // Add event listeners for sorting templates
    const sortableHeaders = document.querySelectorAll('#templates-table th.sortable');
    sortableHeaders.forEach(header => {
      header.addEventListener('click', () => {
        const sortField = header.dataset.sort;
        if (sortField) {
          this.sortTemplates(sortField);
        }
      });
    });
  }

  /**
   * Save current email template with name
   */
  saveEmailTemplate() {
    const saveModal = document.getElementById('save-template-modal');
    
    if (saveModal) {
      // Load existing templates for the dropdown
      this.loadTemplatesForDropdown();
      saveModal.style.display = 'block';
    }
  }
  
  /**
   * Load existing templates for the dropdown in the save modal
   */
  loadTemplatesForDropdown() {
    const templateSelect = document.getElementById('template-select');
    if (!templateSelect) return;
    
    // Clear existing options except the first one
    while (templateSelect.options.length > 1) {
      templateSelect.remove(1);
    }
    
    // Get templates from storage
    chrome.storage.local.get(['emailTemplates'], (result) => {
      const templates = result.emailTemplates || [];
      
      if (templates.length === 0) return;
      
      // Sort templates by name for easier selection
      templates.sort((a, b) => a.name.localeCompare(b.name));
      
      // Add options for each template
      templates.forEach(template => {
        const option = document.createElement('option');
        option.value = template.id;
        option.textContent = template.name;
        templateSelect.appendChild(option);
      });
      
      // Add event listener to update name when a template is selected
      templateSelect.addEventListener('change', () => {
        const selectedId = templateSelect.value;
        if (!selectedId) {
          // Clear fields if "Create new template" is selected
          document.getElementById('template-name').value = '';
          return;
        }
        
        // Find the selected template
        const selectedTemplate = templates.find(t => t.id === selectedId);
        if (selectedTemplate) {
          document.getElementById('template-name').value = selectedTemplate.name;
        }
      });
    });
  }
  
  /**
   * Confirm saving a template
   */
  confirmSaveTemplate() {
    const nameInput = document.getElementById('template-name');
    const templateSelect = document.getElementById('template-select');
    
    const name = nameInput.value.trim();
    const selectedTemplateId = templateSelect ? templateSelect.value : '';
    
    if (!name) {
      this.commonUtils.showStatus('Please enter a name for the template', 'error');
      return;
    }
    
    // Get current subject and body
    const subject = this.options.getSubject();
    const body = this.options.getBody();
    
    if (!subject && !body) {
      this.commonUtils.showStatus('No content to save. Please enter a subject or body.', 'error');
      return;
    }
    
    // Save to storage
    chrome.storage.local.get(['emailTemplates'], (result) => {
      const templates = result.emailTemplates || [];
      
      if (selectedTemplateId) {
        // Update existing template
        const existingTemplateIndex = templates.findIndex(t => t.id === selectedTemplateId);
        
        if (existingTemplateIndex !== -1) {
          // Update the existing template
          templates[existingTemplateIndex] = {
            ...templates[existingTemplateIndex],
            name: name,
            subject: subject,
            body: body,
            lastModified: new Date().toISOString()
          };
          
          chrome.storage.local.set({ emailTemplates: templates }, () => {
            this.commonUtils.showStatus(`Template "${name}" updated successfully`, 'success');
            document.getElementById('save-template-modal').style.display = 'none';
            nameInput.value = '';
            templateSelect.value = '';
          });
          return;
        }
      }
      
      // Create new template
      const newTemplate = {
        id: Date.now().toString(),
        name: name,
        subject: subject,
        body: body,
        dateCreated: new Date().toISOString()
      };
      
      templates.push(newTemplate);
      
      chrome.storage.local.set({ emailTemplates: templates }, () => {
        this.commonUtils.showStatus(`Template "${name}" saved successfully`, 'success');
        document.getElementById('save-template-modal').style.display = 'none';
        nameInput.value = '';
        if (templateSelect) templateSelect.value = '';
      });
    });
  }

  /**
   * Show the templates manager modal
   */
  showTemplatesManager() {
    const modal = document.getElementById('email-templates-modal');
    if (!modal) return;
    
    // Load and display templates
    this.loadTemplates();
    modal.style.display = 'block';
  }
  
  /**
   * Load templates from storage and display in the templates manager
   */
  loadTemplates() {
    const templatesList = document.getElementById('templates-list');
    const noTemplatesMessage = document.getElementById('no-templates-message');
    
    if (!templatesList || !noTemplatesMessage) return;
    
    // Clear the list
    templatesList.innerHTML = '';
    
    // Get templates from storage
    chrome.storage.local.get(['emailTemplates'], (result) => {
      const templates = result.emailTemplates || [];
      
      if (templates.length === 0) {
        noTemplatesMessage.style.display = 'block';
        return;
      }
      
      noTemplatesMessage.style.display = 'none';
      
      // Sort templates based on current sort settings
      this.sortTemplatesArray(templates);
      
      // Display templates
      templates.forEach(template => {
        const row = document.createElement('tr');
        
        // Checkbox cell
        const checkboxCell = document.createElement('td');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.dataset.id = template.id;
        checkboxCell.appendChild(checkbox);
        row.appendChild(checkboxCell);
        
        // Name cell
        const nameCell = document.createElement('td');
        nameCell.textContent = template.name;
        row.appendChild(nameCell);
        
        // Subject cell
        const subjectCell = document.createElement('td');
        subjectCell.textContent = template.subject || '(No subject)';
        row.appendChild(subjectCell);
        
        // Date cell
        const dateCell = document.createElement('td');
        const date = new Date(template.dateCreated);
        dateCell.textContent = date.toLocaleString();
        row.appendChild(dateCell);
        
        // Actions cell
        const actionsCell = document.createElement('td');
        
        // Preview button
        const previewBtn = document.createElement('button');
        previewBtn.textContent = 'Preview';
        previewBtn.className = 'action-button';
        previewBtn.addEventListener('click', () => this.previewTemplate(template));
        
        // Apply button
        const applyBtn = document.createElement('button');
        applyBtn.textContent = 'Apply';
        applyBtn.className = 'action-button';
        applyBtn.addEventListener('click', () => this.applyTemplate(template));
        
        actionsCell.appendChild(previewBtn);
        actionsCell.appendChild(applyBtn);
        row.appendChild(actionsCell);
        
        templatesList.appendChild(row);
      });
    });
  }
  
  /**
   * Sort templates based on the specified field
   * @param {string} field - Field to sort by
   */
  sortTemplates(field) {
    // Toggle direction if sorting by the same field
    if (this.currentSort.field === field) {
      this.currentSort.direction = this.currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
      this.currentSort.field = field;
      this.currentSort.direction = 'asc';
    }
    
    // Update sort indicators
    const sortableHeaders = document.querySelectorAll('#templates-table th.sortable');
    sortableHeaders.forEach(header => {
      const sortIcon = header.querySelector('.sort-icon');
      if (header.dataset.sort === field) {
        sortIcon.innerHTML = this.currentSort.direction === 'asc' ? '&#8593;' : '&#8595;';
      } else {
        sortIcon.innerHTML = '&#8597;';
      }
    });
    
    // Reload templates with new sort
    this.loadTemplates();
  }
  
  /**
   * Sort an array of templates based on current sort settings
   * @param {Array} templates - Array of templates to sort
   */
  sortTemplatesArray(templates) {
    const { field, direction } = this.currentSort;
    
    templates.sort((a, b) => {
      let valueA = a[field];
      let valueB = b[field];
      
      // Handle date fields
      if (field.includes('date')) {
        valueA = new Date(valueA).getTime();
        valueB = new Date(valueB).getTime();
      } else {
        // Convert to lowercase strings for comparison
        valueA = String(valueA || '').toLowerCase();
        valueB = String(valueB || '').toLowerCase();
      }
      
      // Compare values
      if (valueA < valueB) return direction === 'asc' ? -1 : 1;
      if (valueA > valueB) return direction === 'asc' ? 1 : -1;
      return 0;
    });
  }
  
  /**
   * Preview a template
   * @param {Object} template - Template to preview
   */
  previewTemplate(template) {
    const modal = document.getElementById('preview-template-modal');
    if (!modal) return;
    
    // Store the current template for apply button
    this.currentTemplate = template;
    
    // Update preview content
    document.getElementById('preview-template-name').textContent = template.name;
    
    const date = new Date(template.dateCreated);
    document.getElementById('preview-template-date').textContent = date.toLocaleString();
    
    document.getElementById('preview-template-subject').value = template.subject || '';
    document.getElementById('preview-template-body').innerHTML = template.body || '';
    
    // Show the modal
    modal.style.display = 'block';
  }
  
  /**
   * Apply the current template being previewed
   */
  applyCurrentTemplate() {
    if (this.currentTemplate) {
      this.applyTemplate(this.currentTemplate);
      document.getElementById('preview-template-modal').style.display = 'none';
    }
  }
  
  /**
   * Apply a template to the current email
   * @param {Object} template - Template to apply
   */
  applyTemplate(template) {
    // Set subject and body
    this.options.setSubject(template.subject || '');
    this.options.setBody(template.body || '');
    
    // Close the templates manager modal
    const modal = document.getElementById('email-templates-modal');
    if (modal) modal.style.display = 'none';
    
    this.commonUtils.showStatus(`Template "${template.name}" applied successfully`, 'success');
  }
  
  /**
   * Delete selected templates
   */
  deleteSelectedTemplates() {
    const checkboxes = document.querySelectorAll('#templates-list input[type="checkbox"]:checked');
    
    if (checkboxes.length === 0) {
      this.commonUtils.showStatus('No templates selected for deletion', 'error');
      return;
    }
    
    const selectedIds = Array.from(checkboxes).map(checkbox => checkbox.dataset.id);
    
    // Confirm deletion
    if (!confirm(`Are you sure you want to delete ${selectedIds.length} template(s)?`)) {
      return;
    }
    
    // Delete from storage
    chrome.storage.local.get(['emailTemplates'], (result) => {
      let templates = result.emailTemplates || [];
      
      // Filter out the templates to delete
      templates = templates.filter(template => !selectedIds.includes(template.id));
      
      chrome.storage.local.set({ emailTemplates: templates }, () => {
        this.commonUtils.showStatus(`${selectedIds.length} template(s) deleted successfully`, 'success');
        this.loadTemplates();
      });
    });
  }
}

// Export the class
window.EmailTemplatesManager = new EmailTemplatesManager();
