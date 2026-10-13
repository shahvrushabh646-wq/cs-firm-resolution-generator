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

  const bodyOpen = documentXml.indexOf('>', documentXml.indexOf('<w:body')) + 1;
  const bodyClose = documentXml.lastIndexOf('</w:body>');
  if (bodyClose < bodyOpen) {
    throw new Error('Invalid DOCX letterhead package: document body boundaries are invalid.');
  }

  // Keep section properties as the final child of w:body, as required by Word.
  const sectionStart = documentXml.lastIndexOf('<w:sectPr', bodyClose);
  const insertAt = sectionStart >= bodyOpen ? sectionStart : bodyClose;
  let existingBodyContent = documentXml.slice(bodyOpen, insertAt);

  // Trim only trailing empty Word paragraphs, which otherwise consume page space.
  const isEmptyParagraph = (xml) => {
    // Keep any paragraph that contains text, breaks, images, tabs, fields or links.
    return !/<w:(?:t|tab|br|cr|drawing|pict|object|fldChar|instrText|footnoteReference|endnoteReference|commentReference|hyperlink|bookmarkStart|proofErr)\b/.test(
      xml.replace(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/g, '')
         .replace(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/g, '')
    );
  };
  const emptyTail = /<w:p\b[^>]*>[\s\S]*?<\/w:p>\s*$/;
  let tail;
  while ((tail = existingBodyContent.match(emptyTail)) && isEmptyParagraph(tail[0])) {
    existingBodyContent = existingBodyContent.slice(0, tail.index);
  }

  documentXml = documentXml.slice(0, bodyOpen)
    + existingBodyContent
    + processedResolution
    + documentXml.slice(insertAt);

  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}
