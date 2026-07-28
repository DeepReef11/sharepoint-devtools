/**
 * Clipboard Utility Tests
 * Validates clipboard copy functionality and visual feedback
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { copyToClipboard, showCopyFeedback, setupCopyButtons } from '../src/utils/clipboard';

describe('Clipboard Utilities', () => {
  describe('copyToClipboard', () => {
    beforeEach(() => {
      // Mock modern clipboard API
      const mockWriteText = jest.fn<(text: string) => Promise<void>>();
      mockWriteText.mockResolvedValue(undefined);
      Object.assign(navigator, {
        clipboard: {
          writeText: mockWriteText,
        },
      });
    });

    it('should copy text using modern clipboard API', async () => {
      const text = 'test text';
      const result = await copyToClipboard(text);

      expect(result).toBe(true);
      expect(navigator.clipboard.writeText).toHaveBeenCalledWith(text);
    });

    it('should handle clipboard API errors gracefully', async () => {
      const mockWriteText = jest.fn<(text: string) => Promise<void>>();
      mockWriteText.mockRejectedValue(new Error('Permission denied'));
      Object.assign(navigator, {
        clipboard: {
          writeText: mockWriteText,
        },
      });

      const result = await copyToClipboard('test');
      expect(result).toBe(false);
    });

    it('should fallback to execCommand when clipboard API is not available', async () => {
      // Remove modern clipboard API
      Object.defineProperty(navigator, 'clipboard', {
        value: undefined,
        writable: true,
      });

      const mockExecCommand = jest.fn<(command: string) => boolean>();
      mockExecCommand.mockReturnValue(true);
      document.execCommand = mockExecCommand as any;

      const result = await copyToClipboard('test text');
      expect(result).toBe(true);
    });

    it('should handle empty string', async () => {
      const result = await copyToClipboard('');
      expect(result).toBe(true);
      expect(navigator.clipboard.writeText).toHaveBeenCalledWith('');
    });
  });

  describe('showCopyFeedback', () => {
    let button: HTMLElement;

    beforeEach(() => {
      button = document.createElement('button');
      button.innerHTML = 'Copy';
      button.title = 'Copy to clipboard';
      document.body.appendChild(button);
      jest.useFakeTimers();
    });

    afterEach(() => {
      document.body.removeChild(button);
      jest.useRealTimers();
    });

    it('should show success feedback', () => {
      showCopyFeedback(button, true);

      expect(button.innerHTML).toBe('✓');
      expect(button.title).toBe('Copied!');
      expect(button.classList.contains('copied')).toBe(true);
    });

    it('should show failure feedback', () => {
      showCopyFeedback(button, false);

      expect(button.innerHTML).toBe('✗');
      expect(button.title).toBe('Failed to copy');
      expect(button.classList.contains('copy-failed')).toBe(true);
    });

    it('should restore original content after timeout', () => {
      const originalContent = button.innerHTML;
      const originalTitle = button.title;

      showCopyFeedback(button, true);
      jest.advanceTimersByTime(1500);

      expect(button.innerHTML).toBe(originalContent);
      expect(button.title).toBe(originalTitle);
      expect(button.classList.contains('copied')).toBe(false);
    });
  });

  describe('setupCopyButtons', () => {
    let container: HTMLElement;
    let copyButton: HTMLElement;

    beforeEach(() => {
      container = document.createElement('div');
      copyButton = document.createElement('button');
      copyButton.classList.add('sp-btn-copy');
      copyButton.setAttribute('data-value', 'test-value');
      container.appendChild(copyButton);
      document.body.appendChild(container);

      // Mock copyToClipboard
      const mockWriteText = jest.fn<(text: string) => Promise<void>>();
      mockWriteText.mockResolvedValue(undefined);
      Object.assign(navigator, {
        clipboard: {
          writeText: mockWriteText,
        },
      });
    });

    afterEach(() => {
      document.body.removeChild(container);
    });

    it('should setup click event listeners on copy buttons', async () => {
      setupCopyButtons(container);

      copyButton.click();
      // Wait for async operation
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(navigator.clipboard.writeText).toHaveBeenCalledWith('test-value');
    });

    it('should handle clicks on nested elements within copy buttons', async () => {
      const icon = document.createElement('span');
      copyButton.appendChild(icon);

      setupCopyButtons(container);

      icon.click();
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(navigator.clipboard.writeText).toHaveBeenCalledWith('test-value');
    });

    it('should not copy if data-value is missing', async () => {
      copyButton.removeAttribute('data-value');

      setupCopyButtons(container);
      copyButton.click();
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(navigator.clipboard.writeText).not.toHaveBeenCalled();
    });

    it('should prevent default and stop propagation on copy button clicks', async () => {
      setupCopyButtons(container);

      const event = new MouseEvent('click', { bubbles: true, cancelable: true });
      const preventDefaultSpy = jest.spyOn(event, 'preventDefault');
      const stopPropagationSpy = jest.spyOn(event, 'stopPropagation');

      copyButton.dispatchEvent(event);
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(preventDefaultSpy).toHaveBeenCalled();
      expect(stopPropagationSpy).toHaveBeenCalled();
    });
  });
});
