import JSZip from 'jszip';
import { base64ToArrayBuffer } from './fileHelpers.js';

function escapeXml(unsafe) {
  return String(unsafe ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function replacePlaceholders(text, placeholders = {}) {
  let result = String(text ?? '');
  if (placeholders.COMPANY_NAME) {
    result = result.replace(/WAM INDIA PRIVATE LIMITED/gi, String(placeholders.COMPANY_NAME));
    result = result.replace(/\[COMPANY_NAME\]/gi, String(placeholders.COMPANY_NAME));
  }
  for (const [key, value] of Object.entries(placeholders)) {
    const token = '[' + key + ']';
    result = result.split(token).join(value === undefined || value === null || value === '' ? token : String(value));
  }
  return result;
}

function buildCompactResolutionXml(resolutionText, placeholders = {}) {
  const processedText = replacePlaceholders(resolutionText, placeholders);
  const lines = processedText.split(/\r?\n/);
  let xml = '<w:p><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr></w:p>';

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      xml += '<w:p><w:pPr><w:spacing w:after="0" w:line="220" w:lineRule="auto"/></w:pPr></w:p>';
      continue;
    }
    xml += '<w:p><w:pPr><w:spacing w:before="0" w:after="80" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">' + escapeXml(trimmed) + '</w:t></w:r></w:p>';
  }
  return xml;
}

function replaceXmlPlaceholder(xml, key, value) {
  return xml.split('[' + key + ']').join(escapeXml(value));
}

/**
 * Supports both the current app's legacy positional call with ready-made
 * WordprocessingML, and the object-style API with plain-text resolution content.
 */
export async function generateMergedResolutionDocx(
  letterheadOrOptions,
  resolutionBodyXmlSnippet,
  placeholderReplacements = {},
) {
  const objectStyle = letterheadOrOptions && typeof letterheadOrOptions === 'object' && !ArrayBuffer.isView(letterheadOrOptions);
  const letterheadBase64 = objectStyle ? letterheadOrOptions.letterheadBase64 : letterheadOrOptions;
  const replacements = objectStyle ? (letterheadOrOptions.placeholders || {}) : (placeholderReplacements || {});
  const resolutionContent = objectStyle ? letterheadOrOptions.resolutionContent : resolutionBodyXmlSnippet;

  if (!letterheadBase64) {
    throw new Error('Letterhead file data is missing. Please select a company with an uploaded Word letterhead.');
  }

  const letterheadBuffer = base64ToArrayBuffer(letterheadBase64);
  const letterheadZip = await JSZip.loadAsync(letterheadBuffer);
  const docFile = letterheadZip.file('word/document.xml');
  if (!docFile) {
    throw new Error('Invalid DOCX letterhead: word/document.xml not found.');
  }

  let documentXml = await docFile.async('string');
  if (replacements.COMPANY_NAME) {
    documentXml = documentXml.replace(/WAM INDIA PRIVATE LIMITED/gi, escapeXml(replacements.COMPANY_NAME));
  }
  for (const [key, value] of Object.entries(replacements)) {
    if (value !== undefined && value !== null && value !== '') {
      documentXml = replaceXmlPlaceholder(documentXml, key, value);
    }
  }

  // Object-style calls accept plain text and build compact paragraphs. The app's
  // existing positional call passes ready-made WordprocessingML; preserve it.
  let resolutionXml;
  if (objectStyle) {
    resolutionXml = buildCompactResolutionXml(resolutionContent, replacements);
  } else {
    resolutionXml = String(resolutionContent || '');
    for (const [key, value] of Object.entries(replacements)) {
      if (value !== undefined && value !== null && value !== '') {
        resolutionXml = replaceXmlPlaceholder(resolutionXml, key, value);
      }
    }
    resolutionXml = resolutionXml
      .replace(/<w:br\b(?=[^>]*\bw:type="page")[^>]*\/\s*>/g, '')
      .replace(/<w:pageBreakBefore\b[^>]*\/\s*>/g, '')
      .replace(/<w:spacing\b([^>]*)\/\s*>/g, (tag, attrs) => {
        const clean = attrs
          .replace(/\s+w:before="[^"]*"/g, '')
          .replace(/\s+w:after="[^"]*"/g, '')
          .replace(/\s+w:line="[^"]*"/g, '')
          .replace(/\s+w:lineRule="[^"]*"/g, '');
        return '<w:spacing' + clean + ' w:before="0" w:after="60" w:line="240" w:lineRule="auto"/>';
      });
  }

  const bodyOpenMatch = documentXml.match(/<w:body\b[^>]*>/);
  const bodyCloseIndex = documentXml.lastIndexOf('</w:body>');
  if (!bodyOpenMatch || bodyCloseIndex === -1 || bodyCloseIndex < bodyOpenMatch.index + bodyOpenMatch[0].length) {
    throw new Error('Invalid Word document structure.');
  }

  const bodyOpenIndex = bodyOpenMatch.index + bodyOpenMatch[0].length;
  const sectionMatch = documentXml.slice(bodyOpenIndex, bodyCloseIndex).match(/<w:sectPr\b[^>]*(?:\/>|>)/g);
  let insertAt = bodyCloseIndex;
  if (sectionMatch?.length) {
    const lastSectionTag = sectionMatch[sectionMatch.length - 1];
    const sectionIndex = documentXml.lastIndexOf(lastSectionTag, bodyCloseIndex);
    if (sectionIndex >= bodyOpenIndex) insertAt = sectionIndex;
  }

  const updatedXml = documentXml.slice(0, insertAt) + resolutionXml + documentXml.slice(insertAt);
  letterheadZip.file('word/document.xml', updatedXml);

  return letterheadZip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}
