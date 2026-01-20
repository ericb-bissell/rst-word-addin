/**
 * RST Word Add-in - Ribbon Commands
 * Handles function commands executed from ribbon buttons
 */

import { convertToRstAsync, ExtractedImage } from '../converter';

interface OoxmlImage {
  name: string;
  base64: string;
  contentType: string;
}

/**
 * Extract images from OOXML package data
 * OOXML contains images as pkg:part elements with pkg:binaryData
 */
function extractImagesFromOoxml(ooxml: string): OoxmlImage[] {
  const images: OoxmlImage[] = [];

  const parser = new DOMParser();
  const doc = parser.parseFromString(ooxml, 'application/xml');

  const parts = doc.getElementsByTagName('pkg:part');

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const name = part.getAttribute('pkg:name') || '';
    const contentType = part.getAttribute('pkg:contentType') || '';

    if (name.includes('/media/') && contentType.startsWith('image/')) {
      const binaryDataElements = part.getElementsByTagName('pkg:binaryData');
      if (binaryDataElements.length > 0) {
        const base64 = binaryDataElements[0].textContent || '';
        if (base64) {
          images.push({
            name: name,
            base64: base64.replace(/\s/g, ''),
            contentType: contentType,
          });
        }
      }
    }
  }

  return images;
}

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

      // Get OOXML to extract ALL images (including those in text boxes/shapes)
      const ooxmlResult = body.getOoxml();

      await context.sync();

      // Extract images from OOXML
      const ooxmlImages = extractImagesFromOoxml(ooxmlResult.value);

      const html = htmlResult.value;
      const result = await convertToRstAsync(html, {
        imageDirectory: 'images/',
      });

      // Merge OOXML image data with parsed images
      for (let i = 0; i < result.images.length && i < ooxmlImages.length; i++) {
        const img = result.images[i];
        const data = ooxmlImages[i];
        if (!img.base64Data && data.base64) {
          img.base64Data = data.base64;
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
