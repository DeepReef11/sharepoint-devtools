/**
 * Inspector Modal UI Component
 * Provides a modal overlay for the SharePoint Object Inspector
 */

import './inspector-modal.css';

export class InspectorModal {
  private overlay: HTMLElement | null = null;
  private modal: HTMLElement | null = null;
  private isVisible = false;
  private onClose?: () => void;

  constructor() {
    this.createModal();
    this.setupEventListeners();
  }

  /**
   * Creates the modal DOM structure
   */
  private createModal(): void {
    // Create overlay
    this.overlay = document.createElement('div');
    this.overlay.id = 'sp-inspector-overlay';
    this.overlay.className = 'sp-inspector-overlay';

    // Create modal container
    this.modal = document.createElement('div');
    this.modal.id = 'sp-inspector-modal';
    this.modal.className = 'sp-inspector-modal';

    // Create modal header
    const header = document.createElement('div');
    header.className = 'sp-inspector-header';

    const title = document.createElement('h2');
    title.textContent = 'SharePoint Object Inspector';
    title.className = 'sp-inspector-title';

    const closeButton = document.createElement('button');
    closeButton.className = 'sp-inspector-close';
    closeButton.innerHTML = '&times;';
    closeButton.setAttribute('aria-label', 'Close');
    closeButton.onclick = () => this.hide();

    header.appendChild(title);
    header.appendChild(closeButton);

    // Create modal body
    const body = document.createElement('div');
    body.className = 'sp-inspector-body';
    body.id = 'sp-inspector-body';

    // Create inspector section
    const inspectorSection = document.createElement('div');
    inspectorSection.className = 'sp-inspector-section';
    inspectorSection.id = 'sp-inspector-content';

    body.appendChild(inspectorSection);

    // Assemble modal
    this.modal.appendChild(header);
    this.modal.appendChild(body);
    this.overlay.appendChild(this.modal);
  }

  /**
   * Sets up event listeners for the modal
   */
  private setupEventListeners(): void {
    // Close on overlay click
    this.overlay?.addEventListener('click', (e) => {
      if (e.target === this.overlay) {
        this.hide();
      }
    });

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
      if (!this.isVisible) return;

      if (e.key === 'Escape') {
        this.hide();
      } else if (e.key === '/' && !(e.target as HTMLElement)?.matches('input')) {
        e.preventDefault();
        const filterInput = this.modal?.querySelector('.sp-field-filter') as HTMLInputElement;
        if (filterInput) {
          filterInput.focus();
          filterInput.select();
        }
      }
    });
  }

  /**
   * Shows the modal
   */
  public show(): void {
    if (!this.overlay || !this.modal) {
      return;
    }

    if (!document.body.contains(this.overlay)) {
      document.body.appendChild(this.overlay);
    }

    // Trigger reflow for animation
    void this.overlay.offsetHeight;

    this.overlay.classList.add('visible');
    this.modal.classList.add('visible');
    this.isVisible = true;

    // Focus handled by ObjectInspector component
  }

  /**
   * Hides the modal
   */
  public hide(): void {
    if (!this.overlay || !this.modal) {
      return;
    }

    this.overlay.classList.remove('visible');
    this.modal.classList.remove('visible');
    this.isVisible = false;

    // Call onClose callback if provided
    if (this.onClose) {
      this.onClose();
    }

    // Remove from DOM after animation
    setTimeout(() => {
      if (this.overlay && document.body.contains(this.overlay)) {
        document.body.removeChild(this.overlay);
      }
    }, 300);
  }

  /**
   * Toggles the modal visibility
   */
  public toggle(): void {
    if (this.isVisible) {
      this.hide();
    } else {
      this.show();
    }
  }

  /**
   * Checks if the modal is currently visible
   */
  public isOpen(): boolean {
    return this.isVisible;
  }

  /**
   * Gets the inspector section element
   */
  public getInspectorElement(): HTMLElement | null {
    return document.getElementById('sp-inspector-content');
  }

  /**
   * Sets a callback to be called when the modal is closed
   */
  public setOnClose(callback: () => void): void {
    this.onClose = callback;
  }

  /**
   * Destroys the modal and removes it from the DOM
   */
  public destroy(): void {
    if (this.overlay && document.body.contains(this.overlay)) {
      document.body.removeChild(this.overlay);
    }
    this.overlay = null;
    this.modal = null;
    this.isVisible = false;
  }
}
