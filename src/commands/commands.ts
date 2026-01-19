/**
 * RST Word Add-in - Ribbon Commands
 * Handles function commands executed from ribbon buttons
 */

import { convertToRst, ExtractedImage } from '../converter';

// Initialize Office
Office.onReady(() => {
  // Office is ready
});

/**
 * Export document to RST (ribbon command)
 * @param event - The Office event
 */
async function exportToRst(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const htmlResult = body.getHtml();
      await context.sync();

      const html = htmlResult.value;
      const result = convertToRst(html, {
        imageDirectory: 'images/',
      });

      // Export based on whether we have images
      if (result.images.length > 0) {
        await exportAsZip(result.rst, result.images);
      } else {
        downloadRstFile(result.rst);
      }
    });
  } catch (error) {
    console.error('Export error:', error);
  }

  // Signal completion
  event.completed();
}

/**
 * Copy RST to clipboard (ribbon command)
 * @param event - The Office event
 */
async function copyRstToClipboard(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const htmlResult = body.getHtml();
      await context.sync();

      const html = htmlResult.value;
      const result = convertToRst(html);

      await navigator.clipboard.writeText(result.rst);
    });
  } catch (error) {
    console.error('Copy error:', error);
  }

  // Signal completion
  event.completed();
}

/**
 * Download RST file directly
 */
function downloadRstFile(rst: string): void {
  const blob = new Blob([rst], { type: 'text/x-rst' });
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
async function exportAsZip(rst: string, images: ExtractedImage[]): Promise<void> {
  // Dynamically import JSZip
  const JSZip = (await import('jszip')).default;

  const zip = new JSZip();

  // Add RST file
  zip.file('document.rst', rst);

  // Create images folder and add images
  const imagesFolder = zip.folder('images');
  if (imagesFolder) {
    for (const image of images) {
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

// Register commands with Office
Office.actions.associate('exportToRst', exportToRst);
Office.actions.associate('copyRstToClipboard', copyRstToClipboard);
