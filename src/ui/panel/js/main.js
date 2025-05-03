// main.js - Main entry point for panel functionality

console.log('main.js loaded');

document.addEventListener('DOMContentLoaded', function() {
  console.log('DOMContentLoaded event fired');
  
  // Check if all required objects are available
  console.log('CommonUtils available:', typeof CommonUtils !== 'undefined');
  console.log('EmailTab available:', typeof EmailTab !== 'undefined');
  console.log('EventsTab available:', typeof EventsTab !== 'undefined');
  console.log('BulkEmailTab available:', typeof BulkEmailTab !== 'undefined');
  
  // Show loading indicator
  document.getElementById('status').textContent = 'Loading...'; 
  document.getElementById('status').className = 'status pending';
  
  try {
    console.log('Attempting to initialize tab switching...');
    // Initialize tab switching and load tab content
    const tabsLoaded = CommonUtils.initTabSwitching();
    console.log('Tab switching initialized:', tabsLoaded);
    
    if (tabsLoaded) {
      console.log('Initializing rich text editors...');
      // Initialize rich text editors after tab content is loaded
      CommonUtils.initRichTextEditors();
      
      console.log('Initializing tabs...');
      // Initialize each tab
      EmailTab.init();
      EventsTab.init();
      BulkEmailTab.init();
      console.log('All tabs initialized');
      
      // Clear loading status
      document.getElementById('status').textContent = '';
      document.getElementById('status').className = 'status';
    }
  } catch (error) {
    console.error('Error initializing panel:', error);
    document.getElementById('status').textContent = 'Error loading panel. Please refresh the page.';
    document.getElementById('status').className = 'status error';
  }
});
