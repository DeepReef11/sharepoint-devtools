/**
 * SharePoint DevTools - Background Service Worker
 * Handles keyboard shortcuts and extension lifecycle
 */

console.log('SharePoint DevTools - Background Service Worker initialized');

// Keyboard shortcut handler
chrome.commands.onCommand.addListener(async (command) => {
  if (command === 'toggle-quicknav' || command === 'toggle-lists') {
    console.log(`${command} shortcut triggered`);

    try {
      // Get the active tab
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

      if (!tab.id) {
        console.error('No active tab found');
        return;
      }

      // Check if it's a SharePoint page
      const url = tab.url || '';
      const isSharePoint = url.includes('.sharepoint.com') || url.includes('.sharepoint-df.com');

      if (!isSharePoint) {
        console.log('Not a SharePoint page, ignoring shortcut');
        return;
      }

      // Send appropriate message based on command
      const action = command === 'toggle-quicknav' ? 'toggle-modal' : 'toggle-lists';

      // Send message to content script to toggle modal
      chrome.tabs.sendMessage(tab.id, { action }, (response) => {
        if (chrome.runtime.lastError) {
          console.error('Error sending message:', chrome.runtime.lastError);
        } else {
          console.log('Modal toggle response:', response);
        }
      });
    } catch (error) {
      console.error('Error handling keyboard shortcut:', error);
    }
  }
});

// Extension installation handler
chrome.runtime.onInstalled.addListener(() => {
  console.log('SharePoint DevTools extension installed');
  console.log('Keyboard shortcuts:');
  console.log('  - Ctrl+K (Cmd+K on Mac) or Alt+N: Open QuickNav');
  console.log('  - Alt+O: Open Lists & Libraries');
});
