/**
 * Clipboard Utilities
 * Provides copy-to-clipboard functionality with visual feedback
 */

/**
 * Copies text to clipboard
 */
export async function copyToClipboard(text: string): Promise<boolean> {
  try {
    // Modern clipboard API
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }

    // Fallback for older browsers
    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.style.position = 'fixed';
    textarea.style.left = '-9999px';
    textarea.style.top = '-9999px';
    document.body.appendChild(textarea);
    textarea.select();

    try {
      const successful = document.execCommand('copy');
      document.body.removeChild(textarea);
      return successful;
    } catch {
      document.body.removeChild(textarea);
      return false;
    }
  } catch (error) {
    console.error('Failed to copy to clipboard:', error);
    return false;
  }
}

/**
 * Shows a visual feedback when text is copied
 */
export function showCopyFeedback(button: HTMLElement, success = true): void {
  const originalContent = button.innerHTML;
  const originalTitle = button.title;

  if (success) {
    button.innerHTML = '✓';
    button.title = 'Copied!';
    button.classList.add('copied');
  } else {
    button.innerHTML = '✗';
    button.title = 'Failed to copy';
    button.classList.add('copy-failed');
  }

  setTimeout(() => {
    button.innerHTML = originalContent;
    button.title = originalTitle;
    button.classList.remove('copied', 'copy-failed');
  }, 1500);
}

/**
 * Sets up copy button event listeners on a container
 */
export function setupCopyButtons(container: HTMLElement): void {
  container.addEventListener('click', async (e) => {
    const target = e.target as HTMLElement;

    // Check if the clicked element is a copy button
    if (target.classList.contains('sp-btn-copy') || target.closest('.sp-btn-copy')) {
      const button = target.classList.contains('sp-btn-copy')
        ? target
        : (target.closest('.sp-btn-copy') as HTMLElement);

      if (!button) return;

      const value = button.getAttribute('data-value');
      if (!value) return;

      e.preventDefault();
      e.stopPropagation();

      const success = await copyToClipboard(value);
      showCopyFeedback(button, success);
    }
  });
}
