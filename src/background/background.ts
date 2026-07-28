/**
 * Background service worker for SharePoint DevTools extension
 * Handles message passing, keyboard shortcuts, and coordinates between components
 */

import { isSharePointUrl } from '../utils/sharepoint-clouds';
import {
  MessageType,
  ExtensionMessage,
  MessageResponse,
  SharePointDetectedMessage,
  GetContextMessage,
} from '../types/messages';

/**
 * Extension installation handler
 */
chrome.runtime.onInstalled.addListener((details) => {
  console.log('SharePoint DevTools extension installed:', details.reason);

  if (details.reason === 'install') {
    // First time installation
    console.log('First time installation - setting up defaults');
  } else if (details.reason === 'update') {
    // Extension updated
    console.log('Extension updated to version:', chrome.runtime.getManifest().version);
  }
});

/**
 * Keyboard command handler for Ctrl+K / Cmd+K
 */
chrome.commands.onCommand.addListener((command) => {
  console.log('Command triggered:', command);

  if (command === 'open-quicknav') {
    // Send message to active tab to open QuickNav modal
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      if (tabs[0]?.id) {
        chrome.tabs
          .sendMessage(tabs[0].id, {
            type: MessageType.OPEN_QUICKNAV,
            timestamp: Date.now(),
          })
          .then(() => {
            console.log('QuickNav open command sent to content script');
          })
          .catch((error) => {
            console.error('Error sending QuickNav command:', error);
          });
      }
    });
  }
});

/**
 * Message handler for communication with content scripts
 */
chrome.runtime.onMessage.addListener(
  (
    message: ExtensionMessage,
    sender: chrome.runtime.MessageSender,
    sendResponse: (response: MessageResponse) => void
  ) => {
    console.log('Background received message:', message.type, 'from tab:', sender.tab?.id);

    // Handle different message types
    switch (message.type) {
      case MessageType.SHAREPOINT_DETECTED:
        handleSharePointDetected(message as SharePointDetectedMessage, sender);
        sendResponse({ success: true });
        break;

      case MessageType.GET_CONTEXT:
        handleGetContext(message as GetContextMessage, sender)
          .then((data) => {
            sendResponse({ success: true, data });
          })
          .catch((error) => {
            sendResponse({ success: false, error: error.message });
          });
        // Return true to indicate we'll send response asynchronously
        return true;

      default:
        console.warn('Unknown message type:', message.type);
        sendResponse({ success: false, error: 'Unknown message type' });
    }

    return false;
  }
);

/**
 * Handles SharePoint detection messages from content script
 */
function handleSharePointDetected(
  message: SharePointDetectedMessage,
  sender: chrome.runtime.MessageSender
): void {
  const { data } = message;
  console.log('SharePoint page detected:', {
    url: data.url,
    isSharePoint: data.isSharePoint,
    isModern: data.isModern,
    tabId: sender.tab?.id,
  });

  // Update extension icon/badge based on SharePoint detection
  if (sender.tab?.id) {
    if (data.isSharePoint) {
      chrome.action.setBadgeText({
        text: '✓',
        tabId: sender.tab.id,
      });
      chrome.action.setBadgeBackgroundColor({
        color: '#0078d4', // Microsoft blue
        tabId: sender.tab.id,
      });
      chrome.action.setTitle({
        title: `SharePoint DevTools (${data.isModern ? 'Modern' : 'Classic'})`,
        tabId: sender.tab.id,
      });
    } else {
      chrome.action.setBadgeText({
        text: '',
        tabId: sender.tab.id,
      });
      chrome.action.setTitle({
        title: 'SharePoint DevTools (Not on SharePoint)',
        tabId: sender.tab.id,
      });
    }
  }
}

/**
 * Handles context request messages from content script
 */
async function handleGetContext(
  message: GetContextMessage,
  sender: chrome.runtime.MessageSender
): Promise<{ timestamp: number; tabId?: number }> {
  // In a real implementation, this would fetch additional context
  // For now, we'll return a basic response
  console.log('Context requested from tab:', sender.tab?.id);

  return {
    timestamp: Date.now(),
    tabId: sender.tab?.id,
  };
}

/**
 * Tab update listener to detect navigation to SharePoint pages
 */
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status === 'complete' && tab.url) {
    // Check if it's a SharePoint URL
    const isSharePoint = isSharePointUrl(tab.url);

    if (isSharePoint) {
      console.log('SharePoint page loaded in tab:', tabId);
    }
  }
});

/**
 * Extension icon click handler
 */
chrome.action.onClicked.addListener((tab) => {
  if (tab.id) {
    console.log('Extension icon clicked for tab:', tab.id);

    // Send message to open QuickNav
    chrome.tabs
      .sendMessage(tab.id, {
        type: MessageType.OPEN_QUICKNAV,
        timestamp: Date.now(),
      })
      .then(() => {
        console.log('QuickNav open command sent via icon click');
      })
      .catch((error) => {
        console.error('Error sending QuickNav command via icon click:', error);
      });
  }
});

console.log('SharePoint DevTools background service worker initialized');
