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

function normalizePlaceholders(placeholders = {}) {
  return {
    ...placeholders,
    COMPANY_NAME: placeholders.COMPANY_NAME || 'BIT4ID INDIA PRIVATE LIMITED',
    DIRECTOR_NAME:
      placeholders.DIRECTOR_NAME ||
      placeholders.SIGNATORY_NAME ||
      placeholders.directorName ||
      'xyz',
    SIGNATORY_NAME:
      placeholders.SIGNATORY_NAME ||
      placeholders.DIRECTOR_NAME ||
      placeholders.signatoryName ||
      placeholders.directorName ||
      'xyz',
    DESIGNATION: placeholders.DESIGNATION || placeholders.designation || 'Director',
    DIN: placeholders.DIN ?? placeholders.din ?? '',
  };
}

function replacePlaceholders(text, placeholders = {}) {
  let result = String(text ?? '');
  const plc = normalizePlaceholders(placeholders);
  const company = plc.COMPANY_NAME;
  const director = plc.DIRECTOR_NAME || plc.SIGNATORY_NAME;
  const designation = plc.DESIGNATION;
  const din = plc.DIN;

  result = result.replace(/WAM\s+INDIA\s+PRIVATE\s+LIMITED/gi, String(company));
  result = result.replace(/BIT4ID\s+INDIA\s+PRIVATE\s+LIMITED/gi, String(company));
  result = result.replace(/\[COMPANY_NAME\]/gi, String(company));

  result = result.replace(/Mr\.?\s*Luigi\s*Bordonaro/gi, String(director));
  result = result.replace(/Luigi\s+Bordonaro/gi, String(director));
  result = result.replace(/\[(?:DIRECTOR_NAME|SIGNATORY_NAME)\]/gi, String(director));
  result = result.replace(/\[Signatory\s*name\]/gi, String(director));

  if (din !== '') {
    result = result.replace(/03095710/g, String(din));
    // Escape the brackets: /[DIN]/ would incorrectly match individual characters.
    result = result.replace(/\[DIN\]/gi, String(din));
  }
  result = result.replace(/MANAGING\s+DIRECTOR/gi, String(designation));
  result = result.replace(/\[DESIGNATION\]/gi, String(designation));

  for (const [key, value] of Object.entries(placeholders || {})) {
    if (value === undefined || value === null || value === '') continue;
    result = result.split('[' + key + ']').join(String(value));
  }
  return result;
}

function buildResolutionXmlSnippet(content, placeholders = {}) {
  const processed = replacePlaceholders(content, placeholders);
  const trimmedContent = processed.trim();

  // The legacy caller may already provide formatted WordprocessingML.
  if (/^<(?:w:p|w:table|w:tbl)\b/i.test(trimmedContent)) {
    return trimmedContent;
  }

  const lines = processed.split(/\r?\n/);
  let xml = '<w:p><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr></w:p>';

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      xml += '<w:p><w:pPr><w:spacing w:after="0" w:line="220" w:lineRule="auto"/></w:pPr></w:p>';
      continue;
    }
    xml += '<w:p><w:pPr><w:spacing w:before="0" w:after="80" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">' +
      escapeXml(trimmed) +
      '</w:t></w:r></w:p>';
  }
  return xml;
}

function replaceXmlPlaceholder(xml, key, value) {
  return xml.split('[' + key + ']').join(escapeXml(value));
}

function replaceLetterheadText(xml, placeholders = {}) {
  let result = xml;
  const plc = normalizePlaceholders(placeholders);
  const company = plc.COMPANY_NAME;
  const director = plc.DIRECTOR_NAME || plc.SIGNATORY_NAME;
  const designation = plc.DESIGNATION;
  const din = plc.DIN;

  result = result.replace(/WAM\s+INDIA\s+PRIVATE\s+LIMITED/gi, escapeXml(company));
  result = result.replace(/BIT4ID\s+INDIA\s+PRIVATE\s+LIMITED/gi, escapeXml(company));
  result = result.replace(/Mr\.?\s*Luigi\s*Bordonaro/gi, escapeXml(director));
  result = result.replace(/Luigi\s+Bordonaro/gi, escapeXml(director));
  result = result.replace(/MANAGING\s+DIRECTOR/gi, escapeXml(designation));
  if (din !== '') result = result.replace(/03095710/g, escapeXml(din));

  for (const [key, value] of Object.entries(plc)) {
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

  const letterheadBase64 = objectStyle
    ? letterheadOrOptions.letterheadBase64
    : letterheadOrOptions;
  const replacements = normalizePlaceholders(
    objectStyle
      ? (letterheadOrOptions.placeholders || {})
      : (placeholderReplacements || {}),
  );
  const resolutionContent = objectStyle
    ? (letterheadOrOptions.resolutionContent ?? letterheadOrOptions.resolutionBodyXmlSnippet)
    : resolutionBodyXmlSnippet;

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
  documentXml = replaceLetterheadText(documentXml, replacements);

  let resolutionXml = buildResolutionXmlSnippet(resolutionContent, replacements);
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

  const bodyOpenMatch = documentXml.match(/<w:body\b[^>]*>/);
  const bodyCloseIndex = documentXml.lastIndexOf('</w:body>');
  if (
    !bodyOpenMatch ||
    bodyCloseIndex === -1 ||
    bodyCloseIndex < bodyOpenMatch.index + bodyOpenMatch[0].length
  ) {
    throw new Error('Invalid Word document structure: </w:body> not found.');
  }

  const bodyOpenIndex = bodyOpenMatch.index + bodyOpenMatch[0].length;
  const bodyContent = documentXml.slice(bodyOpenIndex, bodyCloseIndex);
  const sectionMatches = [...bodyContent.matchAll(/<w:sectPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:sectPr>)/g)];
  const insertAt = sectionMatches.length
    ? bodyOpenIndex + sectionMatches[sectionMatches.length - 1].index
    : bodyCloseIndex;

  const updatedXml = documentXml.slice(0, insertAt) + resolutionXml + documentXml.slice(insertAt);
  letterheadZip.file('word/document.xml', updatedXml);

  return letterheadZip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}
