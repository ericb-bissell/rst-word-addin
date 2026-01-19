/**
 * RST Word Add-in - Converter Module
 * Main entry point for document conversion functionality
 */

// Main converter
export {
  convertToRst,
  convertToRstAsync,
  convertWordDocument,
  validateRst,
  getConversionStats,
  extractImages,
  previewConversion,
  ConversionOptions,
  ConversionResult,
} from './word-to-rst';

// HTML Parser
export {
  parseWordHtml,
  resetImageCounter,
  ParsedDocument,
  DocumentMetadata,
} from './html-parser';

// RST Formatter
export {
  formatElement,
  formatDocument,
  escapeRstText,
  createLabel,
  createRef,
  createLink,
  createSubstitution,
  createComment,
  formatInlineCode,
  formatBold,
  formatItalic,
  formatCodeBlock,
  formatField,
  formatDefinition,
  FormatterOptions,
} from './rst-formatter';

// Types
export * from './types';

// Directives
export * from './directives';
