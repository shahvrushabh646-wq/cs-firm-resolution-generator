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
 * Replace a literal placeholder in XML text without treating the key as regex.
 * This is intended for placeholders that are contiguous in the XML source.
 */
function replaceXmlPlaceholder(xml, key, value) {
  const token = '[' + key + ']';
  return xml.split(token).join(escapeXml(value));
}

/**
 * Append resolution WordprocessingML to an uploaded letterhead DOCX.
 *
 * The uploaded DOCX remains the base package so existing styles, images,
 * relationships, tables, headers and footers are retained. The resolution is
 * inserted before the final section properties, as required by Word.
 *
 * resolutionBodyXmlSnippet must be valid WordprocessingML paragraphs. Build
 * those paragraphs with the desired font/size; this function deliberately
 * does not force a font or globally rewrite the letterhead's formatting.
 */
export async function generateMergedResolutionDocx(
  letterheadBase64,
  resolutionBodyXmlSnippet,
  placeholderReplacements = {},
) {
  if (!letterheadBase64) {
    throw new Error('Letterhead file data is missing. Please upload the .docx file again.');
  }

  const letterheadBuffer = base64ToArrayBuffer(letterheadBase64);
  const zip = await JSZip.loadAsync(letterheadBuffer);
  const documentFile = zip.file('word/document.xml');

  if (!documentFile) {
    throw new Error('Invalid DOCX letterhead package: word/document.xml not found.');
  }

  let documentXml = await documentFile.async('string');
  const bodyMatch = documentXml.match(/<w:body\b[^>]*>/);
  const bodyClose = documentXml.lastIndexOf('</w:body>');

  if (!bodyMatch || bodyClose < bodyMatch.index + bodyMatch[0].length) {
    throw new Error('Invalid DOCX letterhead package: document body boundaries are invalid.');
  }

  // Replace literal placeholders in the existing document and generated
  // resolution XML. Values are escaped before being inserted into XML.
  for (const [key, value] of Object.entries(placeholderReplacements || {})) {
    if (value === undefined || value === null || String(value) === '') continue;
    documentXml = replaceXmlPlaceholder(documentXml, key, value);
  }

  let processedResolution = String(resolutionBodyXmlSnippet || '');
  for (const [key, value] of Object.entries(placeholderReplacements || {})) {
    if (value === undefined || value === null || String(value) === '') continue;
    processedResolution = replaceXmlPlaceholder(processedResolution, key, value);
  }

  // Compact explicit page breaks and paragraph spacing in the generated resolution
  // only. Do not rewrite the uploaded letterhead's fonts or its paragraph styling.
  processedResolution = processedResolution
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

  const bodyOpen = bodyMatch.index + bodyMatch[0].length;
  const finalBodyClose = documentXml.lastIndexOf('</w:body>');
  if (finalBodyClose < bodyOpen) {
    throw new Error('Invalid DOCX letterhead package: document body boundaries are invalid.');
  }

  // w:sectPr must remain the final child of w:body. Support both self-closing
  // and paired sectPr elements, while leaving their contents untouched.
  const sectionCandidates = [...documentXml.slice(bodyOpen, finalBodyClose).matchAll(/<w:sectPr\b[^>]*(?:\/>|>)/g)];
  let insertAt = finalBodyClose;
  if (sectionCandidates.length) {
    const last = sectionCandidates[sectionCandidates.length - 1];
    insertAt = bodyOpen + last.index;
  }

  let existingBodyContent = documentXml.slice(bodyOpen, insertAt);

  // Remove only trailing empty paragraphs from the existing letterhead.
  // Never remove paragraphs containing text, drawings, fields, breaks,
  // hyperlinks, bookmarks or other visible/semantic content.
  const emptyTail = /<w:p\b[^>]*>[\s\S]*?<\/w:p>\s*$/;
  const isEmptyParagraph = (paragraphXml) => {
    const withoutProperties = paragraphXml
      .replace(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/g, '')
      .replace(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/g, '');
    return !/<w:(?:t|tab|br|cr|drawing|pict|object|fldChar|instrText|footnoteReference|endnoteReference|commentReference|hyperlink|bookmarkStart|proofErr)\b/.test(withoutProperties);
  };

  let tailMatch;
  while ((tailMatch = existingBodyContent.match(emptyTail)) && isEmptyParagraph(tailMatch[0])) {
    existingBodyContent = existingBodyContent.slice(0, tailMatch.index);
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
