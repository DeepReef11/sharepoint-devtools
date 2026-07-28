/**
 * Object Inspector Integration for Content Script
 * Initializes and manages the SharePoint Object Inspector
 */

import { InspectorModal } from '../ui/inspector/InspectorModal';
import { ObjectInspector } from '../ui/object-inspector';
import { setupCopyButtons } from '../utils/clipboard';
import { isSharePointPage } from '../context/sharepoint-context';

export class ObjectInspectorIntegration {
  private modal: InspectorModal | null = null;
  private inspector: ObjectInspector | null = null;

  /**
   * Initialize the object inspector
   */
  init(): void {
    // Only initialize on SharePoint pages
    if (!isSharePointPage()) {
      return;
    }

    // Create modal instance
    this.modal = new InspectorModal();

    // Listen for keyboard shortcut (Alt+I) - supports both left and right Alt
    document.addEventListener('keydown', (e) => {
      if ((e.altKey || e.getModifierState('AltGraph')) && e.key.toLowerCase() === 'i') {
        e.preventDefault();
        this.toggleInspector();
      }
    });

    console.log('Object Inspector initialized (Alt+I to open)');
  }

  /**
   * Toggle the inspector modal
   */
  private toggleInspector(): void {
    if (!this.modal) {
      return;
    }

    if (this.modal.isOpen()) {
      this.modal.hide();
    } else {
      this.modal.show();
      this.loadInspector();
    }
  }

  /**
   * Load the Object Inspector
   */
  private async loadInspector(): Promise<void> {
    const inspectorElement = this.modal?.getInspectorElement();

    if (!inspectorElement) {
      console.error('Inspector element not found');
      return;
    }

    // Create inspector instance if not already created
    if (!this.inspector) {
      this.inspector = new ObjectInspector(inspectorElement);
    }

    // Load SharePoint metadata
    await this.inspector.load();

    // Setup copy button event listeners
    setupCopyButtons(inspectorElement);
  }
}
