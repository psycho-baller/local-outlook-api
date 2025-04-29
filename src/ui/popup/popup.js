document.addEventListener('DOMContentLoaded', function() {
  const openOutlookBtn = document.getElementById('openOutlookBtn');
  
  openOutlookBtn.addEventListener('click', function() {
    chrome.tabs.create({ url: 'https://outlook.office.com' });
    window.close();
  });
});
