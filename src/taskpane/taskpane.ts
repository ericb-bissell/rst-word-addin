/**
 * RST Word Add-in - Taskpane
 * Main entry point for the taskpane UI
 */

import './taskpane.css';
import { convertToRst, ConversionResult, ExtractedImage } from '../converter';

// Version for debugging cache issues
const VERSION = '1.0.9';

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
let currentImages: ExtractedImage[] = [];
let conversionResult: ConversionResult | null = null;
let isLoading: boolean = false;

/**
 * Initialize the add-in when Office is ready
 */
Office.onReady((info) => {
  console.log('Office.onReady called', info);
  try {
    if (info.host === Office.HostType.Word) {
      initializeUI();
      bindEvents();
      setStatus('Ready');
    } else {
      // For testing outside of Word, still initialize UI
      console.log('Running outside Word, initializing anyway for host:', info.host);
      initializeUI();
      bindEvents();
      setStatus('Ready (standalone mode)');
    }
  } catch (error) {
    console.error('Initialization error:', error);
    const errorDiv = document.createElement('pre');
    errorDiv.style.color = 'red';
    errorDiv.textContent = `Initialization Error: ${error}`;
    document.body.prepend(errorDiv);
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
  console.log('handleRefresh called');
  if (isLoading) return;

  isLoading = true;
  showState('loading');
  setStatus('Converting...');

  try {
    console.log('Starting Word.run...');
    await Word.run(async (context) => {
      console.log('Inside Word.run');
      // Get document body
      const body = context.document.body;

      // Get HTML representation
      console.log('Getting HTML...');
      const htmlResult = body.getHtml();

      await context.sync();
      console.log('Got HTML, length:', htmlResult.value?.length);

      const html = htmlResult.value;

      // Convert HTML to RST using the converter module
      console.log('Converting HTML to RST...');
      console.log('HTML preview:', html?.substring(0, 500));

      conversionResult = convertToRst(html, {
        includeMetadata: false,
        addGeneratedComment: false,
        imageDirectory: 'images/',
      });

      console.log('Conversion complete, RST length:', conversionResult.rst?.length);
      console.log('RST preview:', conversionResult.rst?.substring(0, 200));

      currentRst = conversionResult.rst;
      currentImages = conversionResult.images;

      // Update preview - always show debug info at the end
      const elemCount = conversionResult.elements?.length ?? 0;
      const elemTypes = conversionResult.elements?.map(e => e.type).join(', ') || 'none';
      const warnings = conversionResult.warnings.join(', ') || 'none';

      const debugInfo = `

--- DEBUG INFO ---
Version: ${VERSION}
Elements found: ${elemCount}
Element types: ${elemTypes}
Warnings: ${warnings}

--- Raw HTML from Word ---
${html}

--- END DEBUG ---`;

      if (!currentRst) {
        const elemDetails = conversionResult.elements?.map(e => JSON.stringify(e, null, 2)).join('\n\n') || 'none';
        rstPreview.textContent = `(No content converted)
${debugInfo}

--- ELEMENT DETAILS ---
${elemDetails}`;
      } else {
        rstPreview.textContent = currentRst + debugInfo;
      }
      showState('preview');

      // Build status message
      let statusMsg = `Preview updated - ${currentRst.length} characters`;
      if (currentImages.length > 0) {
        statusMsg += `, ${currentImages.length} image${currentImages.length > 1 ? 's' : ''}`;
      }
      if (conversionResult.warnings.length > 0) {
        statusMsg += ` (${conversionResult.warnings.length} warning${conversionResult.warnings.length > 1 ? 's' : ''})`;
        console.warn('Conversion warnings:', conversionResult.warnings);
      }
      setStatus(statusMsg);
    });
  } catch (error) {
    console.error('Error converting document:', error);
    const errorMsg = error instanceof Error ? error.message : String(error);
    console.error('Error details:', errorMsg);
    showError(`Error: ${errorMsg}`);
    // Also show in preview for debugging
    if (rstPreview) {
      rstPreview.textContent = `ERROR: ${errorMsg}\n\nStack: ${error instanceof Error ? error.stack : 'N/A'}`;
      showState('preview');
    }
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
    // If there are images, create a ZIP file
    if (currentImages.length > 0) {
      await exportAsZip();
    } else {
      // No images, just download the RST file
      downloadRstFile();
    }

    showToast('Export complete', 'success');
    setStatus('Export complete');
  } catch (error) {
    console.error('Error exporting:', error);
    showToast('Failed to export', 'error');
    setStatus('Export failed');
  }
}

/**
 * Download RST file directly
 */
function downloadRstFile(): void {
  const blob = new Blob([currentRst], { type: 'text/x-rst' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'document.rst';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Export as ZIP file with RST and images
 */
async function exportAsZip(): Promise<void> {
  // Dynamically import JSZip
  const JSZip = (await import('jszip')).default;

  const zip = new JSZip();

  // Add RST file
  zip.file('document.rst', currentRst);

  // Create images folder and add images
  const imagesFolder = zip.folder('images');
  if (imagesFolder) {
    for (const image of currentImages) {
      if (image.base64Data) {
        // Extract just the filename from the path
        const filename = image.filename.replace('images/', '');
        imagesFolder.file(filename, image.base64Data, { base64: true });
      }
    }
  }

  // Generate ZIP and download
  const content = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(content);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'document.zip';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
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

