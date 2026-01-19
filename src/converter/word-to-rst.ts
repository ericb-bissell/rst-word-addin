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
 * This version handles image extraction that may require async operations.
 *
 * @param html - HTML content from Word
 * @param options - Conversion options
 * @returns Promise resolving to conversion result
 */
export async function convertToRstAsync(
  html: string,
  options: ConversionOptions = {}
): Promise<ConversionResult> {
  // For now, just call the sync version
  // Future enhancement: handle blob URL fetching, etc.
  return convertToRst(html, options);
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
