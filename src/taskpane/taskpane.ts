/**
 * RST Word Add-in - Taskpane
 * Main entry point for the taskpane UI
 */

import './taskpane.css';

// UI Elements
let refreshBtn: HTMLButtonElement;
let copyBtn: HTMLButtonElement;
let exportBtn: HTMLButtonElement;
let helpBtn: HTMLButtonElement;
let retryBtn: HTMLButtonElement;
let loadingState: HTMLElement;
let errorState: HTMLElement;
let emptyState: HTMLElement;
let previewContainer: HTMLElement;
let rstPreview: HTMLPreElement;
let errorMessage: HTMLElement;
let statusText: HTMLElement;
let helpPanel: HTMLElement;
let toast: HTMLElement;
let toastMessage: HTMLElement;

// State
let currentRst: string = '';
let isLoading: boolean = false;

/**
 * Initialize the add-in when Office is ready
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeUI();
    bindEvents();
    setStatus('Ready');
  }
});

/**
 * Initialize UI element references
 */
function initializeUI(): void {
  refreshBtn = document.getElementById('refresh-btn') as HTMLButtonElement;
  copyBtn = document.getElementById('copy-btn') as HTMLButtonElement;
  exportBtn = document.getElementById('export-btn') as HTMLButtonElement;
  helpBtn = document.getElementById('help-btn') as HTMLButtonElement;
  retryBtn = document.getElementById('retry-btn') as HTMLButtonElement;
  loadingState = document.getElementById('loading-state') as HTMLElement;
  errorState = document.getElementById('error-state') as HTMLElement;
  emptyState = document.getElementById('empty-state') as HTMLElement;
  previewContainer = document.getElementById('preview-container') as HTMLElement;
  rstPreview = document.getElementById('rst-preview') as HTMLPreElement;
  errorMessage = document.getElementById('error-message') as HTMLElement;
  statusText = document.getElementById('status-text') as HTMLElement;
  helpPanel = document.getElementById('help-panel') as HTMLElement;
  toast = document.getElementById('toast') as HTMLElement;
  toastMessage = document.getElementById('toast-message') as HTMLElement;
}

/**
 * Bind event listeners
 */
function bindEvents(): void {
  refreshBtn?.addEventListener('click', handleRefresh);
  copyBtn?.addEventListener('click', handleCopy);
  exportBtn?.addEventListener('click', handleExport);
  helpBtn?.addEventListener('click', toggleHelp);
  retryBtn?.addEventListener('click', handleRefresh);

  // Listen for messages from help iframe
  window.addEventListener('message', (event) => {
    if (event.data?.action === 'closeHelp') {
      hideHelp();
    }
  });

  // Keyboard shortcuts
  document.addEventListener('keydown', (event) => {
    if (event.ctrlKey || event.metaKey) {
      switch (event.key.toLowerCase()) {
        case 'r':
          event.preventDefault();
          handleRefresh();
          break;
        case 'c':
          if (currentRst && !window.getSelection()?.toString()) {
            event.preventDefault();
            handleCopy();
          }
          break;
        case 's':
          event.preventDefault();
          handleExport();
          break;
      }
    }
    if (event.key === 'Escape' && helpPanel.style.display !== 'none') {
      hideHelp();
    }
  });
}

/**
 * Show a specific UI state
 */
function showState(state: 'loading' | 'error' | 'empty' | 'preview'): void {
  loadingState.style.display = state === 'loading' ? 'flex' : 'none';
  errorState.style.display = state === 'error' ? 'flex' : 'none';
  emptyState.style.display = state === 'empty' ? 'flex' : 'none';
  previewContainer.style.display = state === 'preview' ? 'block' : 'none';

  // Disable buttons during loading
  const disabled = state === 'loading';
  refreshBtn.disabled = disabled;
  copyBtn.disabled = disabled || state !== 'preview';
  exportBtn.disabled = disabled || state !== 'preview';
}

/**
 * Set status bar text
 */
function setStatus(text: string): void {
  if (statusText) {
    statusText.textContent = text;
  }
}

/**
 * Show toast notification
 */
function showToast(message: string, type: 'success' | 'error' | 'info' = 'info'): void {
  toast.className = `toast toast-${type}`;
  toastMessage.textContent = message;
  toast.style.display = 'block';

  setTimeout(() => {
    toast.style.display = 'none';
  }, 3000);
}

/**
 * Show error state
 */
function showError(message: string): void {
  errorMessage.textContent = message;
  showState('error');
  setStatus('Error');
}

/**
 * Handle refresh button click
 */
async function handleRefresh(): Promise<void> {
  if (isLoading) return;

  isLoading = true;
  showState('loading');
  setStatus('Converting...');

  try {
    await Word.run(async (context) => {
      // Get document body
      const body = context.document.body;

      // Get HTML representation
      const htmlResult = body.getHtml();

      await context.sync();

      const html = htmlResult.value;

      // Convert HTML to RST (placeholder - actual conversion will be implemented)
      currentRst = convertHtmlToRst(html);

      // Update preview
      rstPreview.textContent = currentRst;
      showState('preview');
      setStatus(`Preview updated - ${currentRst.length} characters`);
    });
  } catch (error) {
    console.error('Error converting document:', error);
    showError(error instanceof Error ? error.message : 'Failed to convert document');
  } finally {
    isLoading = false;
  }
}

/**
 * Handle copy button click
 */
async function handleCopy(): Promise<void> {
  if (!currentRst) {
    showToast('No content to copy', 'error');
    return;
  }

  try {
    await navigator.clipboard.writeText(currentRst);
    showToast('RST copied to clipboard', 'success');
    setStatus('Copied to clipboard');
  } catch (error) {
    console.error('Error copying to clipboard:', error);
    showToast('Failed to copy to clipboard', 'error');
  }
}

/**
 * Handle export button click
 */
async function handleExport(): Promise<void> {
  if (!currentRst) {
    showToast('No content to export', 'error');
    return;
  }

  setStatus('Preparing export...');

  try {
    // For now, just download the RST file
    // Full ZIP export with images will be implemented later
    const blob = new Blob([currentRst], { type: 'text/x-rst' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'document.rst';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('RST file downloaded', 'success');
    setStatus('Export complete');
  } catch (error) {
    console.error('Error exporting:', error);
    showToast('Failed to export', 'error');
    setStatus('Export failed');
  }
}

/**
 * Toggle help panel
 */
function toggleHelp(): void {
  if (helpPanel.style.display === 'none') {
    showHelp();
  } else {
    hideHelp();
  }
}

/**
 * Show help panel
 */
function showHelp(): void {
  helpPanel.style.display = 'block';
}

/**
 * Hide help panel
 */
function hideHelp(): void {
  helpPanel.style.display = 'none';
}

/**
 * Convert HTML to RST (placeholder implementation)
 * This will be replaced with the full converter module
 */
function convertHtmlToRst(html: string): string {
  // Basic placeholder conversion
  // The actual converter will be implemented in the converter module

  // Create a temporary DOM element to parse HTML
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  let rst = '';

  // Process body content
  const body = doc.body;
  if (body) {
    rst = processNode(body);
  }

  return rst.trim();
}

/**
 * Process a DOM node and convert to RST (basic implementation)
 */
function processNode(node: Node): string {
  let result = '';

  node.childNodes.forEach((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      result += child.textContent || '';
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      const element = child as HTMLElement;
      const tagName = element.tagName.toLowerCase();

      switch (tagName) {
        case 'h1':
          const h1Text = element.textContent || '';
          const h1Line = '='.repeat(h1Text.length);
          result += `\n${h1Line}\n${h1Text}\n${h1Line}\n\n`;
          break;

        case 'h2':
          const h2Text = element.textContent || '';
          result += `\n${h2Text}\n${'='.repeat(h2Text.length)}\n\n`;
          break;

        case 'h3':
          const h3Text = element.textContent || '';
          result += `\n${h3Text}\n${'-'.repeat(h3Text.length)}\n\n`;
          break;

        case 'h4':
          const h4Text = element.textContent || '';
          result += `\n${h4Text}\n${'~'.repeat(h4Text.length)}\n\n`;
          break;

        case 'h5':
          const h5Text = element.textContent || '';
          result += `\n${h5Text}\n${'^'.repeat(h5Text.length)}\n\n`;
          break;

        case 'h6':
          const h6Text = element.textContent || '';
          result += `\n${h6Text}\n${'"'.repeat(h6Text.length)}\n\n`;
          break;

        case 'p':
          result += `${processNode(element)}\n\n`;
          break;

        case 'strong':
        case 'b':
          result += `**${element.textContent}**`;
          break;

        case 'em':
        case 'i':
          result += `*${element.textContent}*`;
          break;

        case 'code':
          result += `\`\`${element.textContent}\`\``;
          break;

        case 'a':
          const href = element.getAttribute('href');
          const linkText = element.textContent;
          if (href) {
            result += `\`${linkText} <${href}>\`__`;
          } else {
            result += linkText;
          }
          break;

        case 'ul':
          element.querySelectorAll(':scope > li').forEach((li) => {
            result += `- ${li.textContent?.trim()}\n`;
          });
          result += '\n';
          break;

        case 'ol':
          let num = 1;
          element.querySelectorAll(':scope > li').forEach((li) => {
            result += `${num}. ${li.textContent?.trim()}\n`;
            num++;
          });
          result += '\n';
          break;

        case 'br':
          result += '\n';
          break;

        case 'div':
        case 'span':
          result += processNode(element);
          break;

        default:
          result += processNode(element);
          break;
      }
    }
  });

  return result;
}
