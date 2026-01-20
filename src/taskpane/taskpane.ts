/**
 * RST Word Add-in - Taskpane
 * Main entry point for the taskpane UI
 */

import './taskpane.css';
import { convertToRstAsync, ConversionResult, ExtractedImage } from '../converter';

interface OoxmlImage {
  name: string;
  base64: string;
  contentType: string;
}

// Version for debugging cache issues
const VERSION = '1.0.21';

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
let debugContent: HTMLPreElement;
let copyDebugBtn: HTMLButtonElement;

// State
let currentRst: string = '';
let currentImages: ExtractedImage[] = [];
let currentDebugInfo: string = '';
let currentOoxml: string = '';
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
  debugContent = document.getElementById('debug-content') as HTMLPreElement;
  copyDebugBtn = document.getElementById('copy-debug-btn') as HTMLButtonElement;
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
  copyDebugBtn?.addEventListener('click', handleCopyDebug);

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
 * Extract images from OOXML package data
 * OOXML contains images as pkg:part elements with pkg:binaryData
 */
function extractImagesFromOoxml(ooxml: string): OoxmlImage[] {
  const images: OoxmlImage[] = [];

  // Parse the OOXML string
  const parser = new DOMParser();
  const doc = parser.parseFromString(ooxml, 'application/xml');

  // Find all pkg:part elements (namespace-aware)
  const parts = doc.getElementsByTagName('pkg:part');

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const name = part.getAttribute('pkg:name') || '';
    const contentType = part.getAttribute('pkg:contentType') || '';

    // Check if this is an image part (in /word/media/ folder)
    if (name.includes('/media/') && contentType.startsWith('image/')) {
      // Get the binary data
      const binaryDataElements = part.getElementsByTagName('pkg:binaryData');
      if (binaryDataElements.length > 0) {
        const base64 = binaryDataElements[0].textContent || '';
        if (base64) {
          images.push({
            name: name,
            base64: base64.replace(/\s/g, ''), // Remove any whitespace
            contentType: contentType,
          });
        }
      }
    }
  }

  return images;
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

      // Get OOXML to extract ALL images (including those in text boxes/shapes)
      console.log('Getting OOXML...');
      const ooxmlResult = body.getOoxml();

      await context.sync();
      console.log('Got HTML, length:', htmlResult.value?.length);
      console.log('Got OOXML, length:', ooxmlResult.value?.length);

      // Extract images from OOXML (gets ALL images including those in text boxes/shapes)
      const ooxmlImages = extractImagesFromOoxml(ooxmlResult.value);
      console.log('Found', ooxmlImages.length, 'images in OOXML');

      // Store OOXML for export
      currentOoxml = ooxmlResult.value;
      for (let i = 0; i < ooxmlImages.length; i++) {
        const img = ooxmlImages[i];
        console.log(`  OOXML image ${i + 1}: ${img.name}, type=${img.contentType}, base64 length=${img.base64.length}`);
      }

      const html = htmlResult.value;

      // Convert HTML to RST using the converter module
      console.log('Converting HTML to RST...');
      console.log('HTML preview:', html?.substring(0, 500));

      conversionResult = await convertToRstAsync(html, {
        includeMetadata: false,
        addGeneratedComment: false,
        imageDirectory: 'images/',
      });

      console.log('Conversion complete, RST length:', conversionResult.rst?.length);
      console.log('RST preview:', conversionResult.rst?.substring(0, 200));

      // Merge OOXML image data with parsed images
      console.log(`Merging images: ${conversionResult.images.length} parsed, ${ooxmlImages.length} from OOXML`);
      for (let i = 0; i < conversionResult.images.length; i++) {
        const img = conversionResult.images[i];
        console.log(`Parsed image ${i + 1}: filename=${img.filename}, has base64=${!!img.base64Data}, base64 length=${img.base64Data?.length || 0}`);
      }

      // Match by index (OOXML images appear in document order)
      for (let i = 0; i < conversionResult.images.length && i < ooxmlImages.length; i++) {
        const img = conversionResult.images[i];
        const data = ooxmlImages[i];
        console.log(`Matching image ${i + 1}: parsed has data=${!!img.base64Data}, OOXML has data=${!!data.base64}`);
        if (!img.base64Data && data.base64) {
          img.base64Data = data.base64;
          console.log(`Filled image ${i + 1} with OOXML base64 data, length: ${data.base64.length}`);
        } else if (img.base64Data) {
          console.log(`Image ${i + 1} already has base64 data, length: ${img.base64Data.length}`);
        } else {
          console.log(`Image ${i + 1}: No OOXML data available to merge`);
        }
      }

      // Final check
      console.log('Final image status:');
      for (let i = 0; i < conversionResult.images.length; i++) {
        const img = conversionResult.images[i];
        console.log(`  Image ${i + 1}: ${img.filename}, base64 length=${img.base64Data?.length || 0}`);
      }

      currentRst = conversionResult.rst;
      currentImages = conversionResult.images;

      // Build debug info separately
      const elemCount = conversionResult.elements?.length ?? 0;
      const elemTypes = conversionResult.elements?.map(e => e.type).join(', ') || 'none';
      const warnings = conversionResult.warnings.join(', ') || 'none';
      const elemDetails = conversionResult.elements?.map(e => JSON.stringify(e, null, 2)).join('\n\n') || 'none';

      // Build image debug info
      const imageDebug = conversionResult.images.map((img, i) =>
        `  ${i + 1}. ${img.filename}: base64=${img.base64Data?.length || 0} bytes, format=${img.format}`
      ).join('\n') || '  (none)';

      // Build OOXML image debug
      const ooxmlImageDebug = ooxmlImages.map((img, i) =>
        `  ${i + 1}. ${img.name}: ${img.contentType}, ${img.base64.length} bytes`
      ).join('\n') || '  (none)';

      currentDebugInfo = `Version: ${VERSION}
Elements found: ${elemCount}
Element types: ${elemTypes}
Warnings: ${warnings}

--- IMAGE STATUS ---
Parsed images: ${conversionResult.images.length}
OOXML images: ${ooxmlImages.length}
${ooxmlImageDebug}

Final export images:
${imageDebug}

--- ELEMENT DETAILS ---
${elemDetails}

--- Raw HTML from Word ---
${html}`;

      // Update RST preview (clean, without debug info)
      if (!currentRst) {
        rstPreview.textContent = '(No content converted)';
      } else {
        rstPreview.textContent = currentRst;
      }

      // Update debug panel
      if (debugContent) {
        debugContent.textContent = currentDebugInfo;
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
 * Handle copy debug info button click
 */
async function handleCopyDebug(event: Event): Promise<void> {
  event.stopPropagation(); // Prevent toggle of details element

  if (!currentDebugInfo) {
    showToast('No debug info to copy', 'error');
    return;
  }

  try {
    await navigator.clipboard.writeText(currentDebugInfo);
    showToast('Debug info copied to clipboard', 'success');
  } catch (error) {
    console.error('Error copying debug info:', error);
    showToast('Failed to copy debug info', 'error');
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
      const skippedCount = await exportAsZip();
      if (skippedCount > 0) {
        showToast(`Export complete (${skippedCount} image${skippedCount > 1 ? 's' : ''} could not be exported - Shapes are not supported)`, 'info');
        setStatus(`Export complete (${skippedCount} image${skippedCount > 1 ? 's' : ''} skipped)`);
      } else {
        showToast('Export complete', 'success');
        setStatus('Export complete');
      }
    } else {
      // No images, just download the RST file
      downloadRstFile();
      showToast('Export complete', 'success');
      setStatus('Export complete');
    }
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
 * @returns Number of images that couldn't be exported
 */
async function exportAsZip(): Promise<number> {
  console.log('=== exportAsZip START ===');
  console.log(`currentImages count: ${currentImages.length}`);

  // Dynamically import JSZip
  const JSZip = (await import('jszip')).default;

  const zip = new JSZip();

  // Add RST file
  zip.file('document.rst', currentRst);
  console.log('Added document.rst to ZIP');

  // Create images folder and add images
  const imagesFolder = zip.folder('images');
  let addedCount = 0;
  let skippedCount = 0;

  if (imagesFolder) {
    for (let i = 0; i < currentImages.length; i++) {
      const image = currentImages[i];
      console.log(`Export image ${i + 1}: filename=${image.filename}, base64Data length=${image.base64Data?.length || 0}`);
      if (image.base64Data && image.base64Data.length > 0) {
        // Extract just the filename from the path
        const filename = image.filename.replace('images/', '');
        console.log(`  Adding to ZIP: ${filename}, data length: ${image.base64Data.length}`);
        imagesFolder.file(filename, image.base64Data, { base64: true });
        addedCount++;
      } else {
        // Image couldn't be exported - save OOXML for debugging
        const filename = image.filename.replace('images/', '');
        const ooxmlFilename = filename.replace(/\.[^.]+$/, '.ooxml');
        console.log(`  SKIPPED: No base64Data for ${image.filename}, saving ${ooxmlFilename}`);
        if (currentOoxml) {
          imagesFolder.file(ooxmlFilename, currentOoxml);
          console.log(`  Added ${ooxmlFilename} to ZIP for debugging`);
        }
        skippedCount++;
      }
    }
  }

  console.log(`ZIP summary: ${addedCount} images added, ${skippedCount} skipped`);
  console.log('=== exportAsZip END ===');

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

  return skippedCount;
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

