import JSZip from 'jszip';
import { base64ToArrayBuffer } from './fileHelpers.js';

const escapeXml = (value) => String(value ?? '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&apos;');

const escapeRegExp = (value) => String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Append resolution WordprocessingML to an uploaded letterhead DOCX.
 * The original DOCX package remains the base, retaining its styles,
 * relationships, headers, footers, and media files.
 */
export async function generateMergedResolutionDocx(
  letterheadBase64,
  resolutionBodyXmlSnippet,
  placeholderReplacements = {},
) {
  const letterheadBuffer = base64ToArrayBuffer(letterheadBase64);
  const zip = await JSZip.loadAsync(letterheadBuffer);
  const documentFile = zip.file('word/document.xml');

  if (!documentFile) {
    throw new Error('Invalid DOCX letterhead package: word/document.xml not found.');
  }

  let documentXml = await documentFile.async('string');
  if (!/<w:body\b[^>]*>[\s\S]*<\/w:body>/.test(documentXml)) {
    throw new Error('Invalid DOCX letterhead package: document body not found.');
  }

  let processedResolution = String(resolutionBodyXmlSnippet || '');
  for (const [key, value] of Object.entries(placeholderReplacements || {})) {
    const tokenPattern = new RegExp('\\[' + escapeRegExp(key) + '\\]', 'g');
    const safeValue = escapeXml(value || '');
    processedResolution = processedResolution.replace(tokenPattern, () => safeValue);
  }

  const bodyClose = documentXml.lastIndexOf('</w:body>');
  const bodyOpen = documentXml.indexOf('>', documentXml.indexOf('<w:body')) + 1;
  if (bodyClose < bodyOpen) {
    throw new Error('Invalid DOCX letterhead package: document body boundaries are invalid.');
  }

  // Keep section properties as the final child of w:body, as required by Word.
  const sectionStart = documentXml.lastIndexOf('<w:sectPr', bodyClose);
  const insertAt = sectionStart >= bodyOpen ? sectionStart : bodyClose;
  documentXml = documentXml.slice(0, insertAt)
    + processedResolution
    + documentXml.slice(insertAt);

  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}
