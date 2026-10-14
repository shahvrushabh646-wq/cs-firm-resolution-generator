import JSZip from 'jszip';
import { base64ToArrayBuffer } from './fileHelpers.js';

function escapeXml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function replacePlaceholders(text, placeholders = {}) {
  let result = String(text ?? '');
  const company = placeholders.COMPANY_NAME;
  const director = placeholders.DIRECTOR_NAME || placeholders.SIGNATORY_NAME;
  const designation = placeholders.DESIGNATION;
  const din = placeholders.DIN;

  if (company) {
    result = result.replace(/WAM INDIA PRIVATE LIMITED/gi, String(company));
    result = result.replace(/BIT4ID INDIA PRIVATE LIMITED/gi, String(company));
    result = result.replace(/\[COMPANY_NAME\]/gi, String(company));
  }
  if (director) {
    result = result.replace(/Mr\.?\s*Luigi\s*Bordonaro/gi, String(director));
    result = result.replace(/Luigi\s*Bordonaro/gi, String(director));
    result = result.replace(/\[(?:DIRECTOR_NAME|SIGNATORY_NAME)\]/gi, String(director));
  }
  if (din !== undefined && din !== null && din !== '') {
    result = result.replace(/03095710/g, String(din));
    result = result.replace(/\[DIN\]/gi, String(din));
  }
  if (designation !== undefined && designation !== null && designation !== '') {
    result = result.replace(/Managing\s*Director/gi, String(designation));
    result = result.replace(/\[DESIGNATION\]/gi, String(designation));
  }

  // Literal token replacement avoids constructing fragile dynamic regular expressions.
  for (const [key, value] of Object.entries(placeholders || {})) {
    if (value === undefined || value === null || value === '') continue;
    result = result.split('[' + key + ']').join(String(value));
  }
  return result;
}

function buildCompactResolutionXml(resolutionText, placeholders = {}) {
  const lines = replacePlaceholders(resolutionText, placeholders).split(/\r?\n/);
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

function replaceLetterheadText(xml, placeholders = {}) {
  let result = xml;
  const company = placeholders.COMPANY_NAME;
  const director = placeholders.DIRECTOR_NAME || placeholders.SIGNATORY_NAME;
  const designation = placeholders.DESIGNATION;
  const din = placeholders.DIN;

  if (company) {
    result = result.replace(/WAM INDIA PRIVATE LIMITED/gi, escapeXml(company));
    result = result.replace(/BIT4ID INDIA PRIVATE LIMITED/gi, escapeXml(company));
  }
  if (director) {
    result = result.replace(/Mr\.?\s*Luigi\s*Bordonaro/gi, escapeXml(director));
    result = result.replace(/Luigi\s*Bordonaro/gi, escapeXml(director));
  }
  if (din) result = result.replace(/03095710/g, escapeXml(din));
  if (designation) result = result.replace(/Managing\s*Director/gi, escapeXml(designation));

  for (const [key, value] of Object.entries(placeholders || {})) {
    if (value !== undefined && value !== null && value !== '') {
      result = replaceXmlPlaceholder(result, key, value);
    }
  }
  return result;
}

/**
 * Merge resolution wording into the uploaded letterhead as ONE DOCX.
 * Supports both the legacy positional call and the object-style API.
 */
export async function generateMergedResolutionDocx(
  letterheadOrOptions,
  resolutionBodyXmlSnippet,
  placeholderReplacements = {},
) {
  const objectStyle =
    letterheadOrOptions !== null &&
    typeof letterheadOrOptions === 'object' &&
    !ArrayBuffer.isView(letterheadOrOptions) &&
    !(letterheadOrOptions instanceof ArrayBuffer);

  const letterheadBase64 = objectStyle ? letterheadOrOptions.letterheadBase64 : letterheadOrOptions;
  const replacements = objectStyle
    ? (letterheadOrOptions.placeholders || {})
    : (placeholderReplacements || {});
  const resolutionContent = objectStyle
    ? letterheadOrOptions.resolutionContent
    : resolutionBodyXmlSnippet;

  if (!letterheadBase64) {
    throw new Error('Letterhead file data is missing. Please select a company with an uploaded Word letterhead.');
  }

  const letterheadBuffer = base64ToArrayBuffer(letterheadBase64);
  const letterheadZip = await JSZip.loadAsync(letterheadBuffer);
  const docFile = letterheadZip.file('word/document.xml');
  if (!docFile) throw new Error('Invalid DOCX letterhead: word/document.xml not found.');

  let documentXml = await docFile.async('string');
  documentXml = replaceLetterheadText(documentXml, replacements);

  let resolutionXml;
  if (objectStyle) {
    resolutionXml = buildCompactResolutionXml(resolutionContent, replacements);
  } else {
    // The legacy caller supplies WordprocessingML; retain its formatting.
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
  if (
    !bodyOpenMatch ||
    bodyCloseIndex === -1 ||
    bodyCloseIndex < bodyOpenMatch.index + bodyOpenMatch[0].length
  ) {
    throw new Error('Invalid Word document structure.');
  }

  // Keep section properties as the final child of w:body.
  const bodyOpenIndex = bodyOpenMatch.index + bodyOpenMatch[0].length;
  const bodyContent = documentXml.slice(bodyOpenIndex, bodyCloseIndex);
  const sectionMatches = [...bodyContent.matchAll(/<w:sectPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:sectPr>)/g)];
  let insertAt = bodyCloseIndex;
  if (sectionMatches.length) {
    insertAt = bodyOpenIndex + sectionMatches[sectionMatches.length - 1].index;
  }

  const updatedXml = documentXml.slice(0, insertAt) + resolutionXml + documentXml.slice(insertAt);
  letterheadZip.file('word/document.xml', updatedXml);

  return letterheadZip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}
