/**
 * Content script for SharePoint DevTools extension
 * Runs on SharePoint pages to detect context, handle keyboard shortcuts, and show modal
 */

import { MessageType, ExtensionMessage } from '../types/messages';
import { SharePointContext, SharePointVersion } from '../types/sharepoint';
import { getSharePointContext } from '../utils/sharepoint-detector';

/**
 * Content script initialization
 */
class SharePointQuickNav {
  private context: SharePointContext | null = null;
  private modalOpen: boolean = false;
  private initialized: boolean = false;

  constructor() {
    this.init();
  }

  /**
   * Initialize the content script
   */
  private init(): void {
    if (this.initialized) {
      return;
    }

    console.log('SharePoint DevTools content script initializing...');

    // Detect SharePoint context
    this.detectContext();

    // Set up keyboard event listener
    this.setupKeyboardListener();

    // Set up message listener for background communication
    this.setupMessageListener();

    this.initialized = true;
    console.log('SharePoint DevTools content script initialized');
  }

  /**
   * Detect SharePoint context and notify background
   */
  private detectContext(): void {
    const currentUrl = window.location.href;
    this.context = getSharePointContext(currentUrl);

    console.log('SharePoint context detected:', this.context);

    // Notify background service worker
    if (this.context.isSharePoint) {
      this.sendMessage({
        type: MessageType.SHAREPOINT_DETECTED,
        timestamp: Date.now(),
        data: {
          url: currentUrl,
          isSharePoint: this.context.isSharePoint,
          isModern: this.context.version === SharePointVersion.MODERN,
        },
      });
    }
  }

  /**
   * Set up keyboard event listener for Ctrl+K / Cmd+K
   */
  private setupKeyboardListener(): void {
    document.addEventListener('keydown', (event: KeyboardEvent) => {
      // Check for Ctrl+K (Windows/Linux) or Cmd+K (Mac)
      const isModifierPressed = event.ctrlKey || event.metaKey;
      const isKKey = event.key === 'k' || event.key === 'K';

      if (isModifierPressed && isKKey && !event.shiftKey && !event.altKey) {
        // Prevent default browser behavior
        event.preventDefault();
        event.stopPropagation();

        console.log('QuickNav keyboard shortcut triggered');
        this.openQuickNav();
      }

      // Also handle Escape to close modal
      if (event.key === 'Escape' && this.modalOpen) {
        event.preventDefault();
        this.closeQuickNav();
      }
    });

    console.log('Keyboard listener set up (Ctrl+K / Cmd+K)');
  }

  /**
   * Set up message listener for background communication
   */
  private setupMessageListener(): void {
    chrome.runtime.onMessage.addListener(
      (
        message: ExtensionMessage,
        sender: chrome.runtime.MessageSender,
        sendResponse: (response: {
          success: boolean;
          data?: SharePointContext;
          error?: string;
        }) => void
      ) => {
        console.log('Content script received message:', message.type);

        switch (message.type) {
          case MessageType.OPEN_QUICKNAV:
            this.openQuickNav();
            sendResponse({ success: true });
            break;

          case MessageType.GET_CONTEXT:
            sendResponse({
              success: true,
              data: this.context ?? undefined,
            });
            break;

          default:
            sendResponse({ success: false, error: 'Unknown message type' });
        }

        return false;
      }
    );

    console.log('Message listener set up');
  }

  /**
   * Open the QuickNav modal
   */
  private openQuickNav(): void {
    if (this.modalOpen) {
      console.log('QuickNav modal already open');
      return;
    }

    if (!this.context?.isSharePoint) {
      console.warn('Not on a SharePoint page, cannot open QuickNav');
      this.showNotification('QuickNav is only available on SharePoint pages');
      return;
    }

    console.log('Opening QuickNav modal...');
    this.modalOpen = true;

    // TODO: Create and show the actual modal UI (Issue #6 - Modal UI Component)
    // For now, just log and show a placeholder
    this.showPlaceholderModal();
  }

  /**
   * Close the QuickNav modal
   */
  private closeQuickNav(): void {
    if (!this.modalOpen) {
      return;
    }

    console.log('Closing QuickNav modal...');
    this.modalOpen = false;

    // Remove the modal from DOM
    const existingModal = document.getElementById('sharepoint-devtools-modal');
    if (existingModal) {
      existingModal.remove();
    }
  }

  /**
   * Show a placeholder modal (temporary until Issue #6 is implemented)
   */
  private showPlaceholderModal(): void {
    // Remove any existing modal
    const existing = document.getElementById('sharepoint-devtools-modal');
    if (existing) {
      existing.remove();
    }

    // Create a simple placeholder modal
    const modal = document.createElement('div');
    modal.id = 'sharepoint-devtools-modal';
    modal.style.cssText = `
      position: fixed;
      top: 20%;
      left: 50%;
      transform: translateX(-50%);
      width: 600px;
      max-width: 90%;
      background: white;
      border-radius: 8px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
      padding: 20px;
      z-index: 999999;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    `;

    modal.innerHTML = `
      <h2 style="margin: 0 0 15px 0; color: #0078d4; font-size: 18px;">
        SharePoint DevTools (Preview)
      </h2>
      <p style="margin: 0 0 10px 0; color: #333; font-size: 14px;">
        <strong>Current Context:</strong>
      </p>
      <ul style="margin: 0 0 15px 0; padding-left: 20px; font-size: 13px; color: #666;">
        <li>SharePoint: ${this.context?.isSharePoint ? 'Yes' : 'No'}</li>
        <li>Version: ${this.context?.version || 'Unknown'}</li>
        <li>Site URL: ${this.context?.siteUrl || 'N/A'}</li>
        <li>Web URL: ${this.context?.webUrl || 'N/A'}</li>
        <li>List ID: ${this.context?.listId || 'N/A'}</li>
        <li>Page Type: ${this.context?.pageType || 'N/A'}</li>
      </ul>
      <p style="margin: 0 0 15px 0; color: #666; font-size: 12px; font-style: italic;">
        Full UI modal will be implemented in Issue #6 (Modal UI Component)
      </p>
      <button id="quicknav-close-btn" style="
        background: #0078d4;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
      ">
        Close (Press ESC)
      </button>
    `;

    // Add overlay
    const overlay = document.createElement('div');
    overlay.id = 'sharepoint-devtools-overlay';
    overlay.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0, 0, 0, 0.5);
      z-index: 999998;
    `;

    // Append to body
    document.body.appendChild(overlay);
    document.body.appendChild(modal);

    // Set up close handlers
    const closeBtn = document.getElementById('quicknav-close-btn');
    if (closeBtn) {
      closeBtn.addEventListener('click', () => this.closeQuickNav());
    }

    overlay.addEventListener('click', () => this.closeQuickNav());
  }

  /**
   * Show a notification to the user
   */
  private showNotification(message: string): void {
    // Simple notification using browser's built-in alert for now
    // TODO: Replace with a nicer notification UI
    console.log('Notification:', message);

    alert(`SharePoint DevTools: ${message}`);
  }

  /**
   * Send a message to the background service worker
   */
  private sendMessage(message: ExtensionMessage): Promise<unknown> {
    return chrome.runtime.sendMessage(message);
  }
}

// Initialize the content script when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    new SharePointQuickNav();
  });
} else {
  // DOM is already ready
  new SharePointQuickNav();
}
