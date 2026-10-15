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
  const company = placeholders.COMPANY_NAME ?? placeholders.companyName ?? '';
  const director =
    placeholders.DIRECTOR_NAME ??
    placeholders.SIGNATORY_NAME ??
    placeholders.directorName ??
    placeholders.signatoryName ??
    '';
  return {
    ...placeholders,
    COMPANY_NAME: company,
    DIRECTOR_NAME: director,
    SIGNATORY_NAME: placeholders.SIGNATORY_NAME ?? director,
    DESIGNATION: placeholders.DESIGNATION ?? placeholders.designation ?? 'Director',
    DIN: placeholders.DIN ?? placeholders.din ?? '',
  };
}

function replaceLiteralTokens(text, placeholders = {}) {
  let result = String(text ?? '');
  const plc = normalizePlaceholders(placeholders);
  const company = String(plc.COMPANY_NAME ?? '');
  const director = String(plc.DIRECTOR_NAME || plc.SIGNATORY_NAME || '');
  const designation = String(plc.DESIGNATION ?? 'Director');
  const din = String(plc.DIN ?? '');

  // Replace legacy sample text only when the relevant master value is supplied.
  if (company) {
    result = result.replace(/WAM\s+INDIA\s+PRIVATE\s+LIMITED/gi, company);
    result = result.replace(/BIT4ID\s+INDIA\s+PRIVATE\s+LIMITED/gi, company);
    result = result.replace(/\[COMPANY_NAME\]/gi, company);
  }
  if (director) {
    result = result.replace(/Mr\.?\s*Luigi\s*Bordonaro/gi, director);
    result = result.replace(/Luigi\s+Bordonaro/gi, director);
    result = result.replace(/\[(?:DIRECTOR_NAME|SIGNATORY_NAME)\]/gi, director);
    result = result.replace(/\[Signatory\s*name\]/gi, director);
  }
  if (din) {
    result = result.replace(/03095710/g, din);
    result = result.replace(/\[DIN\]/gi, din);
  }
  result = result.replace(/MANAGING\s+DIRECTOR/gi, designation);
  result = result.replace(/\[DESIGNATION\]/gi, designation);

  for (const [key, value] of Object.entries(placeholders || {})) {
    if (value === undefined || value === null || value === '') continue;
    // Only treat the key as a literal placeholder, not as a regular expression.
    const token = new RegExp('\\[' + escapeRegExp(key) + '\\]', 'gi');
    result = result.replace(token, String(value));
  }
  return result;
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Replace text inside Word text nodes only. This prevents replacement values
 * containing &, <, or > from corrupting document.xml markup.
 */
function replaceTextNodeContents(xml, placeholders = {}) {
  return xml.replace(/(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g, (whole, open, rawText, close) => {
    // Decode the five XML entities before substitutions, then safely re-escape.
    const decoded = rawText
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
    return open + escapeXml(replaceLiteralTokens(decoded, placeholders)) + close;
  });
}

function buildResolutionXmlSnippet(content, placeholders = {}) {
  const processed = replaceLiteralTokens(content, placeholders);
  const trimmedContent = processed.trim();

  // Content already formatted as WordprocessingML should not be escaped as plain text.
  if (/^<(?:w:p|w:tbl|w:table)\b/i.test(trimmedContent)) {
    // Restrict substitutions to text nodes so replacement values cannot damage XML tags.
    return replaceTextNodeContents(trimmedContent, placeholders);
  }

  const lines = processed.split(/\r?\n/);
  let xml = '';
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

/**
 * Replace legacy names and placeholders within Word text nodes only.
 * XML tags and attributes are deliberately left untouched.
 */
function replaceLetterheadText(xml, placeholders = {}) {
  return replaceTextNodeContents(xml, placeholders);
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
  if (resolutionContent === undefined || resolutionContent === null || String(resolutionContent).trim() === '') {
    throw new Error('Resolution content is empty. Please add the resolution wording before generating the DOCX.');
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
    throw new Error('Invalid Word document structure: <w:body> closing tag not found.');
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
