/**
 * Popup Script for SharePoint DevTools Extension
 */

// Check if current tab is a SharePoint page
async function checkSharePointPage() {
  const statusEl = document.getElementById('status');

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab || !tab.url) {
      statusEl.textContent = '❌ Unable to detect current page';
      statusEl.className = 'status inactive';
      return false;
    }

    const isSharePoint = tab.url.includes('.sharepoint.com') || tab.url.includes('.sharepoint-df.com');

    if (isSharePoint) {
      statusEl.textContent = '✓ Active on SharePoint page';
      statusEl.className = 'status active';
      return true;
    } else {
      statusEl.textContent = '⚠️ Not a SharePoint page';
      statusEl.className = 'status inactive';
      return false;
    }
  } catch (error) {
    console.error('Error checking page:', error);
    statusEl.textContent = '❌ Error checking page';
    statusEl.className = 'status inactive';
    return false;
  }
}

// Open QuickNav modal with retry
async function openQuickNav(retryCount = 0) {
  const MAX_RETRIES = 3;
  const RETRY_DELAY = 500; // ms

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab.id) {
      console.error('No active tab found');
      alert('No active tab found.');
      return;
    }

    // Check if it's a SharePoint page
    if (!tab.url || (!tab.url.includes('.sharepoint.com') && !tab.url.includes('.sharepoint-df.com'))) {
      alert('This extension only works on SharePoint pages.\n\nPlease navigate to a SharePoint site first.');
      return;
    }

    // Send message to content script and wait for response
    chrome.tabs.sendMessage(tab.id, { action: 'toggle-modal' }, (response) => {
      if (chrome.runtime.lastError) {
        console.error('Error sending message (attempt ' + (retryCount + 1) + '):', chrome.runtime.lastError);

        // Retry if connection failed and we haven't exceeded max retries
        if (retryCount < MAX_RETRIES && chrome.runtime.lastError.message.includes('Could not establish connection')) {
          console.log('Retrying in ' + RETRY_DELAY + 'ms...');
          setTimeout(() => openQuickNav(retryCount + 1), RETRY_DELAY);
          return;
        }

        alert('Could not communicate with the page.\n\nTry:\n1. Refresh the SharePoint page (F5)\n2. Reload the extension\n3. Try again\n\nError: ' + chrome.runtime.lastError.message);
        return;
      }

      if (!response || !response.success) {
        console.error('Modal not initialized:', response);
        alert('QuickNav is not ready yet.\n\nTry refreshing the SharePoint page (F5) and try again.\n\nError: ' + (response?.error || 'Unknown error'));
        return;
      }

      // Close popup on success
      window.close();
    });
  } catch (error) {
    console.error('Error opening QuickNav:', error);
    alert('Could not open QuickNav.\n\nError: ' + error.message);
  }
}

// Open options page
function openOptions() {
  chrome.runtime.openOptionsPage();
  window.close();
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  // Check SharePoint status
  checkSharePointPage();

  // Setup button listeners
  document.getElementById('openQuickNav').addEventListener('click', openQuickNav);
  document.getElementById('openOptions').addEventListener('click', openOptions);
});
