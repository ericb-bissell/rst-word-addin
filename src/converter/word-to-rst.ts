/**
 * RST Word Add-in - Main Converter
 * Converts Word document HTML to reStructuredText
 *
 * This is the main entry point for document conversion.
 * It coordinates the HTML parsing and RST formatting pipeline.
 */

import { parseWordHtml, ParsedDocument, DocumentMetadata } from './html-parser';
import { formatDocument, FormatterOptions } from './rst-formatter';
import { ExtractedImage, AnyDocumentElement } from './types';

/**
 * Conversion options
 */
export interface ConversionOptions extends Partial<FormatterOptions> {
  /** Include document metadata as field list */
  includeMetadata?: boolean;
  /** Add generation comment at top */
  addGeneratedComment?: boolean;
  /** Image directory path */
  imageDirectory?: string;
}

/**
 * Conversion result
 */
export interface ConversionResult {
  /** Generated RST content */
  rst: string;
  /** Extracted images */
  images: ExtractedImage[];
  /** Document metadata */
  metadata: DocumentMetadata;
  /** Any warnings during conversion */
  warnings: string[];
  /** Parsed document elements (for debugging) */
  elements?: AnyDocumentElement[];
}

/**
 * Default conversion options
 */
const DEFAULT_CONVERSION_OPTIONS: ConversionOptions = {
  includeMetadata: false,
  addGeneratedComment: false,
  imageDirectory: 'images/',
  lineWidth: 0,
  titleOverline: true,
  indentSize: 3,
};

/**
 * Convert Word HTML to RST
 *
 * @param html - HTML content from Word's getHtml() method
 * @param options - Conversion options
 * @returns Conversion result with RST, images, and metadata
 *
 * @example
 * ```typescript
 * // Get HTML from Word
 * const html = await Word.run(async (context) => {
 *   const body = context.document.body;
 *   body.load('html');
 *   await context.sync();
 *   return body.html;
 * });
 *
 * // Convert to RST
 * const result = convertToRst(html);
 * console.log(result.rst);
 *
 * // Access extracted images
 * for (const image of result.images) {
 *   console.log(`Image: ${image.filename}`);
 * }
 * ```
 */
export function convertToRst(
  html: string,
  options: ConversionOptions = {}
): ConversionResult {
  const opts = { ...DEFAULT_CONVERSION_OPTIONS, ...options };
  const warnings: string[] = [];

  // Parse HTML into document elements
  let parsed: ParsedDocument;
  try {
    parsed = parseWordHtml(html);
  } catch (error) {
    warnings.push(`HTML parsing error: ${error instanceof Error ? error.message : String(error)}`);
    return {
      rst: '',
      images: [],
      metadata: {},
      warnings,
    };
  }

  // Format elements to RST
  const formatterOptions: Partial<FormatterOptions> = {
    lineWidth: opts.lineWidth,
    titleOverline: opts.titleOverline,
    indentSize: opts.indentSize,
    imageDir: opts.imageDirectory || 'images/',
  };

  let rst = formatDocument(parsed.elements, formatterOptions);

  // Add metadata if requested
  if (opts.includeMetadata && hasMetadata(parsed.metadata)) {
    const metadataRst = formatMetadata(parsed.metadata);
    rst = metadataRst + '\n\n' + rst;
  }

  // Add generation comment if requested
  if (opts.addGeneratedComment) {
    const comment = formatGenerationComment();
    rst = comment + '\n\n' + rst;
  }

  return {
    rst,
    images: parsed.images,
    metadata: parsed.metadata,
    warnings,
    elements: parsed.elements,
  };
}

/**
 * Convert Word HTML to RST with async image handling
 *
 * This version handles image extraction that may require async operations,
 * including fetching blob URLs and converting them to base64.
 *
 * @param html - HTML content from Word
 * @param options - Conversion options
 * @returns Promise resolving to conversion result
 */
export async function convertToRstAsync(
  html: string,
  options: ConversionOptions = {}
): Promise<ConversionResult> {
  // First, do the sync conversion
  const result = convertToRst(html, options);

  // Then resolve any blob URLs in images
  await resolveBlobUrls(result.images, result.warnings);

  return result;
}

/**
 * Fetch blob URLs and convert to base64
 *
 * Word Online often provides images as blob: URLs which need to be
 * fetched and converted to base64 for export.
 */
async function resolveBlobUrls(
  images: ExtractedImage[],
  warnings: string[]
): Promise<void> {
  const blobImages = images.filter(
    (img) => !img.base64Data && img.filename
  );

  if (blobImages.length === 0) {
    return;
  }

  // Process images in parallel
  await Promise.all(
    blobImages.map(async (image) => {
      try {
        // Try to get the original src from the image element
        // The src should be stored or we need to find another way
        // For now, we'll try using the canvas approach with the img element
        const base64 = await fetchImageAsBase64(image);
        if (base64) {
          image.base64Data = base64;
        } else {
          warnings.push(`Could not fetch image: ${image.filename}`);
        }
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        warnings.push(`Error fetching image ${image.filename}: ${msg}`);
      }
    })
  );
}

/**
 * Fetch an image and convert to base64
 *
 * Uses fetch for blob URLs, with canvas fallback for other image types.
 */
async function fetchImageAsBase64(image: ExtractedImage): Promise<string | null> {
  const srcUrl = image.srcUrl;
  if (!srcUrl) {
    return null;
  }

  // Handle blob URLs - fetch and convert to base64
  if (srcUrl.startsWith('blob:')) {
    try {
      const response = await fetch(srcUrl);
      if (!response.ok) {
        throw new Error(`Fetch failed: ${response.status}`);
      }
      const blob = await response.blob();
      return await blobToBase64(blob);
    } catch (error) {
      // Blob URL may have expired or be inaccessible
      // Try canvas approach as fallback
      return await loadImageViaCanvas(srcUrl, image.format);
    }
  }

  // For other URLs, try canvas approach
  return await loadImageViaCanvas(srcUrl, image.format);
}

/**
 * Convert a Blob to base64 string (without data URL prefix)
 */
function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const result = reader.result as string;
      // Remove the data URL prefix (e.g., "data:image/png;base64,")
      const base64 = result.split(',')[1];
      resolve(base64 || '');
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

/**
 * Load image via canvas and extract base64
 *
 * This works for images that can be drawn to canvas (same-origin or CORS-enabled).
 */
function loadImageViaCanvas(src: string, format: string): Promise<string | null> {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';

    img.onload = () => {
      try {
        const canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth || img.width;
        canvas.height = img.naturalHeight || img.height;

        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve(null);
          return;
        }

        ctx.drawImage(img, 0, 0);

        // Get base64 data
        const mimeType = format === 'jpg' ? 'image/jpeg' : `image/${format}`;
        const dataUrl = canvas.toDataURL(mimeType);
        const base64 = dataUrl.split(',')[1];
        resolve(base64 || null);
      } catch {
        // Canvas tainted by cross-origin data
        resolve(null);
      }
    };

    img.onerror = () => resolve(null);

    // Set timeout to avoid hanging
    setTimeout(() => resolve(null), 5000);

    img.src = src;
  });
}

/**
 * Convert Word document body directly (Office.js integration)
 *
 * @param context - Word request context
 * @param options - Conversion options
 * @returns Promise resolving to conversion result
 */
export async function convertWordDocument(
  context: Word.RequestContext,
  options: ConversionOptions = {}
): Promise<ConversionResult> {
  const body = context.document.body;

  // Load HTML representation
  const htmlRange = body.getRange();
  htmlRange.load('text');

  // Get OOXML for better structure (optional, for future enhancement)
  // const ooxml = body.getOoxml();

  await context.sync();

  // Get HTML via the clipboard approach or direct method
  // Note: body.getHtml() might need to be implemented via selection
  const html = await getDocumentHtml(context);

  return convertToRst(html, options);
}

/**
 * Get HTML from Word document
 *
 * Word's Office.js doesn't have a direct getHtml() on body,
 * so we need to work with the selection or use alternative approaches.
 */
async function getDocumentHtml(context: Word.RequestContext): Promise<string> {
  // Select entire document
  const body = context.document.body;
  const range = body.getRange('Whole');

  // Get HTML via getHtml() method
  const html = range.getHtml();
  await context.sync();

  return html.value;
}

/**
 * Check if metadata object has any values
 */
function hasMetadata(metadata: DocumentMetadata): boolean {
  return !!(metadata.title || metadata.author || metadata.language);
}

/**
 * Format metadata as RST field list
 */
function formatMetadata(metadata: DocumentMetadata): string {
  const fields: string[] = [];

  if (metadata.title) {
    fields.push(`:title: ${metadata.title}`);
  }
  if (metadata.author) {
    fields.push(`:author: ${metadata.author}`);
  }
  if (metadata.language) {
    fields.push(`:language: ${metadata.language}`);
  }

  return fields.join('\n');
}

/**
 * Format generation comment
 */
function formatGenerationComment(): string {
  const date = new Date().toISOString().split('T')[0];
  return `.. Generated by RST Word Add-in on ${date}`;
}

/**
 * Validate RST output (basic checks)
 *
 * @param rst - RST content to validate
 * @returns Array of validation warnings
 */
export function validateRst(rst: string): string[] {
  const warnings: string[] = [];
  const lines = rst.split('\n');

  let inDirective = false;
  let directiveIndent = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineNum = i + 1;

    // Check for common issues

    // Unmatched inline markup
    const unpairedAsterisks = (line.match(/\*(?!\*)/g) || []).length;
    if (unpairedAsterisks % 2 !== 0) {
      warnings.push(`Line ${lineNum}: Potentially unpaired asterisk (*)`);
    }

    // Check heading underlines match text length
    if (i > 0 && /^[=\-~^"'`]+$/.test(line.trim())) {
      const prevLine = lines[i - 1];
      if (prevLine.trim() && line.trim().length < prevLine.trim().length) {
        warnings.push(`Line ${lineNum}: Heading underline may be too short`);
      }
    }

    // Track directive blocks
    if (line.match(/^\.\.\s+\w+::/)) {
      inDirective = true;
      directiveIndent = 3;
    } else if (inDirective && line.trim() && !line.startsWith(' '.repeat(directiveIndent))) {
      inDirective = false;
    }

    // Check for tabs (RST prefers spaces)
    if (line.includes('\t')) {
      warnings.push(`Line ${lineNum}: Contains tab character (spaces preferred)`);
    }
  }

  return warnings;
}

/**
 * Get statistics about the conversion
 *
 * @param result - Conversion result
 * @returns Statistics object
 */
export function getConversionStats(result: ConversionResult): {
  elementCount: number;
  imageCount: number;
  wordCount: number;
  lineCount: number;
  hasMetadata: boolean;
  warningCount: number;
} {
  const lines = result.rst.split('\n');
  const words = result.rst.split(/\s+/).filter((w) => w.length > 0);

  return {
    elementCount: result.elements?.length || 0,
    imageCount: result.images.length,
    wordCount: words.length,
    lineCount: lines.length,
    hasMetadata: hasMetadata(result.metadata),
    warningCount: result.warnings.length,
  };
}

/**
 * Extract just the images from HTML without full conversion
 *
 * Useful when you only need the images.
 *
 * @param html - HTML content
 * @returns Array of extracted images
 */
export function extractImages(html: string): ExtractedImage[] {
  const parsed = parseWordHtml(html);
  return parsed.images;
}

/**
 * Preview conversion with limited output
 *
 * Useful for showing a preview without processing the entire document.
 *
 * @param html - HTML content
 * @param maxElements - Maximum number of elements to process
 * @returns Preview RST content
 */
export function previewConversion(html: string, maxElements: number = 10): string {
  const parsed = parseWordHtml(html);
  const limitedElements = parsed.elements.slice(0, maxElements);

  let rst = formatDocument(limitedElements, {});

  if (parsed.elements.length > maxElements) {
    rst += '\n\n.. [Preview truncated - ' +
      `${parsed.elements.length - maxElements} more elements]`;
  }

  return rst;
}
