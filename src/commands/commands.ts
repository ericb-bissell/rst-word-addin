/**
 * RST Word Add-in - Ribbon Commands
 * Handles function commands executed from ribbon buttons
 */

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
      const rst = convertHtmlToRst(html);

      // Download RST file
      const blob = new Blob([rst], { type: 'text/x-rst' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'document.rst';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
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
      const rst = convertHtmlToRst(html);

      await navigator.clipboard.writeText(rst);
    });
  } catch (error) {
    console.error('Copy error:', error);
  }

  // Signal completion
  event.completed();
}

/**
 * Basic HTML to RST conversion (placeholder)
 * Full implementation will be in the converter module
 */
function convertHtmlToRst(html: string): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  let rst = '';
  const body = doc.body;

  if (body) {
    rst = processNode(body);
  }

  return rst.trim();
}

/**
 * Process DOM node to RST
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

        default:
          result += processNode(element);
          break;
      }
    }
  });

  return result;
}

// Register commands with Office
Office.actions.associate('exportToRst', exportToRst);
Office.actions.associate('copyRstToClipboard', copyRstToClipboard);
