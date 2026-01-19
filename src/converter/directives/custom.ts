/**
 * RST Word Add-in - Custom Directive Handler
 * Handles Word styles starting with "rst_" and converts them to RST directives
 *
 * Word styles like "rst_note", "rst_warning", "rst_code-block" are converted
 * to their corresponding RST directives.
 */

import { CustomDirective } from '../types';

/**
 * Prefix that identifies custom RST directive styles
 */
const RST_STYLE_PREFIX = 'rst_';

/**
 * Default indentation for directive options and content
 */
const INDENT = '   ';

/**
 * Check if a style name indicates a custom RST directive
 *
 * @param styleName - Word style name
 * @returns True if this is an rst_* style
 */
export function isRstDirectiveStyle(styleName: string): boolean {
  return styleName.toLowerCase().startsWith(RST_STYLE_PREFIX);
}

/**
 * Extract directive name from style name
 *
 * @param styleName - Word style name (e.g., "rst_note", "rst_code-block")
 * @returns Directive name (e.g., "note", "code-block")
 */
export function extractDirectiveName(styleName: string): string {
  return styleName.substring(RST_STYLE_PREFIX.length);
}

/**
 * Parse custom directive content from Word paragraph
 *
 * Content format:
 * ```
 * [argument]           <- Optional, first line in square brackets
 * :option1: value1     <- Optional, lines starting with :name:
 * :option2: value2
 *
 * Body content here    <- Rest becomes directive body
 * More content...
 * ```
 *
 * @param styleName - Word style name (e.g., "rst_note")
 * @param content - Paragraph content
 * @returns Parsed custom directive
 */
export function parseCustomDirective(
  styleName: string,
  content: string
): CustomDirective {
  const directiveName = extractDirectiveName(styleName);
  const lines = content.split('\n');

  let argument: string | undefined;
  const options = new Map<string, string>();
  const bodyLines: string[] = [];

  let inBody = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmedLine = line.trim();

    // Skip empty lines at the beginning
    if (!inBody && trimmedLine === '') {
      continue;
    }

    // Check for argument (first non-empty line in square brackets)
    if (!inBody && !argument && isArgumentLine(trimmedLine)) {
      argument = extractArgument(trimmedLine);
      continue;
    }

    // Check for option line
    if (!inBody && isOptionLine(trimmedLine)) {
      const parsed = parseOptionLine(trimmedLine);
      if (parsed) {
        options.set(parsed.name, parsed.value);
      }
      continue;
    }

    // Everything else is body content
    inBody = true;
    bodyLines.push(line);
  }

  // Trim leading/trailing empty lines from body
  const body = trimBodyContent(bodyLines);

  return {
    name: directiveName,
    argument,
    options,
    content: body,
  };
}

/**
 * Generate RST directive from parsed custom directive
 *
 * @param directive - Parsed custom directive
 * @returns RST directive string
 */
export function generateCustomDirective(directive: CustomDirective): string {
  const lines: string[] = [];

  // Directive declaration with optional argument
  if (directive.argument) {
    lines.push(`.. ${directive.name}:: ${directive.argument}`);
  } else {
    lines.push(`.. ${directive.name}::`);
  }

  // Add options
  for (const [name, value] of directive.options) {
    if (value) {
      lines.push(`${INDENT}:${name}: ${value}`);
    } else {
      // Flag option (no value)
      lines.push(`${INDENT}:${name}:`);
    }
  }

  // Add body content
  if (directive.content) {
    lines.push('');
    const contentLines = directive.content.split('\n');
    for (const line of contentLines) {
      // Indent non-empty lines
      lines.push(line ? `${INDENT}${line}` : '');
    }
  }

  return lines.join('\n');
}

/**
 * Check if a line is an argument line [argument]
 *
 * @param line - Line to check
 * @returns True if line is an argument
 */
function isArgumentLine(line: string): boolean {
  return line.startsWith('[') && line.endsWith(']');
}

/**
 * Extract argument from [argument] line
 *
 * @param line - Argument line
 * @returns Extracted argument
 */
function extractArgument(line: string): string {
  return line.slice(1, -1).trim();
}

/**
 * Check if a line is an option line :name: value
 *
 * @param line - Line to check
 * @returns True if line is an option
 */
function isOptionLine(line: string): boolean {
  return /^:[a-zA-Z][\w-]*:/.test(line);
}

/**
 * Parse option from :name: value line
 *
 * @param line - Option line
 * @returns Parsed option or null
 */
function parseOptionLine(line: string): { name: string; value: string } | null {
  const match = line.match(/^:([a-zA-Z][\w-]*):\s*(.*)$/);
  if (match) {
    return {
      name: match[1],
      value: match[2].trim(),
    };
  }
  return null;
}

/**
 * Trim leading and trailing empty lines from body content
 *
 * @param lines - Body content lines
 * @returns Trimmed body content
 */
function trimBodyContent(lines: string[]): string {
  // Remove leading empty lines
  while (lines.length > 0 && lines[0].trim() === '') {
    lines.shift();
  }

  // Remove trailing empty lines
  while (lines.length > 0 && lines[lines.length - 1].trim() === '') {
    lines.pop();
  }

  return lines.join('\n');
}

/**
 * Convert style name to valid RST directive name
 *
 * Handles various naming conventions:
 * - rst_code-block -> code-block
 * - rst_codeblock -> codeblock (preserved as-is)
 * - rst_CODE_BLOCK -> code-block (normalized)
 *
 * @param styleName - Word style name
 * @returns Valid RST directive name
 */
export function normalizeDirectiveName(styleName: string): string {
  const name = extractDirectiveName(styleName);

  // Replace underscores with hyphens (common convention)
  // But preserve intentional underscores in known directive names
  const knownWithUnderscores = ['code_block']; // Add more as needed

  if (knownWithUnderscores.includes(name.toLowerCase())) {
    return name.toLowerCase();
  }

  return name.toLowerCase().replace(/_/g, '-');
}

/**
 * Get list of common built-in RST/Sphinx directives
 *
 * Used for validation and documentation.
 *
 * @returns Array of known directive names
 */
export function getKnownDirectives(): string[] {
  return [
    // Standard RST admonitions
    'attention',
    'caution',
    'danger',
    'error',
    'hint',
    'important',
    'note',
    'tip',
    'warning',
    'admonition',

    // RST body elements
    'topic',
    'sidebar',
    'line-block',
    'parsed-literal',
    'rubric',
    'epigraph',
    'highlights',
    'pull-quote',
    'compound',
    'container',

    // RST tables
    'table',
    'csv-table',
    'list-table',

    // RST code
    'code',
    'code-block',
    'sourcecode',
    'literalinclude',

    // Sphinx directives
    'toctree',
    'only',
    'index',
    'glossary',
    'productionlist',
    'deprecated',
    'versionadded',
    'versionchanged',
    'seealso',
    'centered',
    'hlist',

    // Sphinx-needs
    'need',
    'req',
    'spec',
    'impl',
    'test',
    'needflow',
    'needtable',
    'needlist',

    // Common extensions
    'todo',
    'todolist',
    'math',
    'graphviz',
    'plantuml',
  ];
}

/**
 * Check if a directive name is a known standard directive
 *
 * @param name - Directive name
 * @returns True if known directive
 */
export function isKnownDirective(name: string): boolean {
  return getKnownDirectives().includes(name.toLowerCase());
}

/**
 * Suggest the correct directive name for common misspellings
 *
 * @param name - Potentially misspelled directive name
 * @returns Suggested correction or null
 */
export function suggestDirectiveName(name: string): string | null {
  const corrections: Record<string, string> = {
    'codeblock': 'code-block',
    'code_block': 'code-block',
    'sourcecode': 'code-block',
    'source-code': 'code-block',
    'warn': 'warning',
    'info': 'note',
    'requirement': 'req',
    'specification': 'spec',
    'implementation': 'impl',
    'testcase': 'test',
    'test-case': 'test',
  };

  return corrections[name.toLowerCase()] || null;
}

/**
 * Merge consecutive paragraphs with the same rst_* style
 *
 * When multiple paragraphs have the same rst_* style, they should
 * be combined into a single directive with multi-paragraph content.
 *
 * @param paragraphs - Array of { style, content } objects
 * @returns Merged paragraphs
 */
export function mergeConsecutiveDirectives(
  paragraphs: Array<{ style: string; content: string }>
): Array<{ style: string; content: string }> {
  const result: Array<{ style: string; content: string }> = [];

  for (const para of paragraphs) {
    const lastResult = result[result.length - 1];

    // Check if this is an rst_* style
    if (!isRstDirectiveStyle(para.style)) {
      result.push(para);
      continue;
    }

    // Check if previous paragraph has the same style
    if (lastResult && lastResult.style === para.style) {
      // Check if this paragraph continues the directive (no argument/options)
      const trimmed = para.content.trim();
      const hasArgumentOrOptions =
        isArgumentLine(trimmed.split('\n')[0]) ||
        isOptionLine(trimmed.split('\n')[0]);

      if (!hasArgumentOrOptions) {
        // Merge with previous
        lastResult.content += '\n\n' + para.content;
        continue;
      }
    }

    result.push(para);
  }

  return result;
}
