/**
 * Template Parser
 * Reads Google Docs templates and replaces placeholders with recipient data
 */

/**
 * Get the template content from Google Doc
 * @param {string} docId - Google Doc ID
 * @returns {string} HTML content of the document
 */
function getTemplateContent(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Get HTML content
    // Note: DocumentApp doesn't directly support HTML export, so we'll get text
    // and preserve basic formatting using approach that works with GmailApp
    return convertBodyToHtml(body);
  } catch (error) {
    throw new Error(`Failed to read template document: ${error.message}`);
  }
}

/**
 * Convert Google Doc body to HTML
 * Preserves basic formatting like bold, italic, links
 * @param {GoogleAppsScript.Document.Body} body - Document body
 * @returns {string} HTML content
 */
function convertBodyToHtml(body) {
  const text = body.getText();
  const numChildren = body.getNumChildren();
  let html = '';

  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = child.asParagraph();
      html += processParagraph(paragraph);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = child.asListItem();
      html += '<li>' + processText(listItem) + '</li>';
    } else if (type === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      html += processTable(table);
    }
  }

  return html || text;
}

/**
 * Process a paragraph element
 * @param {GoogleAppsScript.Document.Paragraph} paragraph
 * @returns {string} HTML paragraph
 */
function processParagraph(paragraph) {
  const text = processText(paragraph);
  const heading = paragraph.getHeading();

  switch (heading) {
    case DocumentApp.ParagraphHeading.HEADING1:
      return `<h1>${text}</h1>`;
    case DocumentApp.ParagraphHeading.HEADING2:
      return `<h2>${text}</h2>`;
    case DocumentApp.ParagraphHeading.HEADING3:
      return `<h3>${text}</h3>`;
    default:
      return `<p>${text}</p>`;
  }
}

/**
 * Process text with formatting
 * Groups consecutive characters with the same formatting to preserve placeholders
 * @param {GoogleAppsScript.Document.Text} element
 * @returns {string} Formatted HTML text
 */
function processText(element) {
  const text = element.getText();
  if (!text || text.trim() === '') {
    return '';
  }

  const textElement = element.editAsText();
  const spans = [];
  let currentSpan = {
    text: '',
    isBold: false,
    isItalic: false,
    linkUrl: null
  };

  // Group consecutive characters with the same formatting
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const isBold = textElement.isBold(i);
    const isItalic = textElement.isItalic(i);
    const linkUrl = textElement.getLinkUrl(i);

    // Check if formatting changed
    const formattingChanged =
      currentSpan.isBold !== isBold ||
      currentSpan.isItalic !== isItalic ||
      currentSpan.linkUrl !== linkUrl;

    if (formattingChanged && currentSpan.text !== '') {
      // Save current span and start a new one
      spans.push({...currentSpan});
      currentSpan = {
        text: char,
        isBold: isBold,
        isItalic: isItalic,
        linkUrl: linkUrl
      };
    } else {
      // Continue current span
      currentSpan.text += char;
      currentSpan.isBold = isBold;
      currentSpan.isItalic = isItalic;
      currentSpan.linkUrl = linkUrl;
    }
  }

  // Add the last span
  if (currentSpan.text !== '') {
    spans.push(currentSpan);
  }

  // Convert spans to HTML
  let html = '';
  for (const span of spans) {
    let formatted = span.text;

    // Apply formatting in the correct order (innermost to outermost)
    if (span.isBold) {
      formatted = `<strong>${formatted}</strong>`;
    }
    if (span.isItalic) {
      formatted = `<em>${formatted}</em>`;
    }
    if (span.linkUrl) {
      formatted = `<a href="${span.linkUrl}">${formatted}</a>`;
    }

    html += formatted;
  }

  return html;
}

/**
 * Process table element (basic support)
 * @param {GoogleAppsScript.Document.Table} table
 * @returns {string} HTML table
 */
function processTable(table) {
  let html = '<table border="1">';
  const numRows = table.getNumRows();

  for (let i = 0; i < numRows; i++) {
    html += '<tr>';
    const row = table.getRow(i);
    const numCells = row.getNumCells();

    for (let j = 0; j < numCells; j++) {
      const cell = row.getCell(j);
      html += `<td>${cell.getText()}</td>`;
    }
    html += '</tr>';
  }

  html += '</table>';
  return html;
}

/**
 * Replace placeholders in template with recipient data
 * @param {string} template - Template content with {{placeholder}} format
 * @param {Object} data - Key-value pairs for replacement
 * @returns {string} Content with placeholders replaced
 */
function replacePlaceholders(template, data) {
  let result = template;

  // Find all placeholders in format {{fieldname}}
  const placeholderRegex = /\{\{(\w+)\}\}/g;
  const matches = template.match(placeholderRegex);

  if (matches) {
    for (const match of matches) {
      // Extract field name (remove {{ and }})
      const fieldName = match.replace(/\{\{|\}\}/g, '');

      // Get value from data, use empty string if not found
      const value = data[fieldName] !== undefined ? data[fieldName] : '';

      // Replace all occurrences of this placeholder
      result = result.replace(new RegExp(match.replace(/\{/g, '\\{').replace(/\}/g, '\\}'), 'g'), value);
    }
  }

  return result;
}

/**
 * Parse template and generate personalized email content
 * @param {string} docId - Template Google Doc ID
 * @param {Object} recipientData - Recipient data for personalization
 * @returns {string} Personalized email content
 */
function parseTemplate(docId, recipientData) {
  try {
    const template = getTemplateContent(docId);
    return replacePlaceholders(template, recipientData);
  } catch (error) {
    throw new Error(`Template parsing failed: ${error.message}`);
  }
}

/**
 * Parse subject line with placeholders
 * @param {string} subjectTemplate - Subject with {{placeholder}} format
 * @param {Object} recipientData - Recipient data
 * @returns {string} Personalized subject line
 */
function parseSubject(subjectTemplate, recipientData) {
  return replacePlaceholders(subjectTemplate, recipientData);
}
