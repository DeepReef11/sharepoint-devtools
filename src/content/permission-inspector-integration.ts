/**
 * Permission Inspector Integration for Content Script
 * Initializes and manages the SharePoint Permission Inspector
 */

import { PermissionInspector } from '../ui/permission-inspector';
import { isSharePointPage } from '../context/sharepoint-context';

export class PermissionInspectorIntegration {
  private inspector: PermissionInspector | null = null;

  /**
   * Initialize the permission inspector
   */
  init(): void {
    if (!isSharePointPage()) {
      return;
    }

    this.inspector = new PermissionInspector();

    // Listen for keyboard shortcut (Alt+P)
    document.addEventListener('keydown', (e) => {
      if ((e.altKey || e.getModifierState('AltGraph')) && e.key.toLowerCase() === 'p') {
        e.preventDefault();
        this.inspector?.toggle();
      }
    });

    console.log('Permission Inspector initialized (Alt+P to open)');
  }
}
