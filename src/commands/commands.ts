/**
 * RST Word Add-in - Ribbon Commands
 * Handles function commands executed from ribbon buttons
 */

import { convertToRstAsync, ExtractedImage } from '../converter';

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

      // Get inline pictures for image extraction
      const inlinePictures = body.inlinePictures;
      inlinePictures.load('items');

      await context.sync();

      // Extract base64 data from each picture using Office.js API
      const pictureData: { base64: string; width: number; height: number }[] = [];
      for (const picture of inlinePictures.items) {
        picture.load(['width', 'height']);
        const base64Result = picture.getBase64ImageSrc();
        await context.sync();
        pictureData.push({
          base64: base64Result.value,
          width: picture.width,
          height: picture.height,
        });
      }

      const html = htmlResult.value;
      const result = await convertToRstAsync(html, {
        imageDirectory: 'images/',
      });

      // Merge Office.js image data with parsed images
      for (let i = 0; i < result.images.length && i < pictureData.length; i++) {
        const img = result.images[i];
        const data = pictureData[i];
        if (!img.base64Data && data.base64) {
          const base64 = data.base64.includes(',')
            ? data.base64.split(',')[1]
            : data.base64;
          img.base64Data = base64;
          img.width = data.width;
          img.height = data.height;
        }
      }

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
 *
 * Note: Clipboard API may not work in the hidden function command context.
 * Falls back to downloading as a text file if clipboard fails.
 */
async function copyRstToClipboard(event: Office.AddinCommands.Event): Promise<void> {
  try {
    let rst = '';

    await Word.run(async (context) => {
      const body = context.document.body;
      const htmlResult = body.getHtml();
      await context.sync();

      const html = htmlResult.value;
      const result = await convertToRstAsync(html);
      rst = result.rst;
    });

    if (!rst) {
      event.completed();
      return;
    }

    // Try clipboard API first
    let clipboardSuccess = false;
    try {
      await navigator.clipboard.writeText(rst);
      clipboardSuccess = true;
    } catch {
      // Clipboard API failed (expected in hidden context)
      console.log('Clipboard API not available, falling back to download');
    }

    // If clipboard failed, download as text file instead
    if (!clipboardSuccess) {
      downloadTextFile(rst, 'document-rst.txt');
    }
  } catch (error) {
    console.error('Copy error:', error);
  }

  // Signal completion
  event.completed();
}

/**
 * Download content as a text file
 */
function downloadTextFile(content: string, filename: string): void {
  const blob = new Blob([content], { type: 'text/plain' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
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
