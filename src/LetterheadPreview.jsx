import React, { useEffect, useRef, useState } from 'react';
import * as docxPreview from 'docx-preview';
import mammoth from 'mammoth/mammoth.browser';
import { base64ToArrayBuffer } from './utils/fileHelpers';

export default function LetterheadPreview({ letterheadData, fileName }) {
  const previewContainerRef = useRef(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [fallbackHtml, setFallbackHtml] = useState(null);

  useEffect(() => {
    let cancelled = false;
    let timeoutId;
    const container = previewContainerRef.current;

    async function renderDocument() {
      if (!letterheadData) {
        setLoading(false);
        setError('No letterhead file uploaded or selected for this company.');
        setFallbackHtml(null);
        if (container) container.replaceChildren();
        return;
      }
      setLoading(true);
      setError(null);
      setFallbackHtml(null);
      if (!container) return;
      container.replaceChildren();

      try {
        const arrayBuffer = typeof letterheadData === 'string'
          ? base64ToArrayBuffer(letterheadData)
          : letterheadData;
        const renderPromise = docxPreview.renderAsync(arrayBuffer, container, undefined, {
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          breakPages: true,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
          experimental: true,
        });
        const timeoutPromise = new Promise((_, reject) => {
          timeoutId = window.setTimeout(() => reject(new Error('Preview rendering timed out.')), 15000);
        });
        await Promise.race([renderPromise, timeoutPromise]);
        if (cancelled) return;
        const hasContent = container.querySelector('section.docx, p, table, img, svg');
        if (!hasContent) throw new Error('The DOCX renderer did not produce visible page content.');
        setLoading(false);
      } catch (renderError) {
        if (cancelled) return;
        console.warn('docx-preview failed; trying Mammoth HTML fallback:', renderError);
        try {
          const arrayBuffer = typeof letterheadData === 'string'
            ? base64ToArrayBuffer(letterheadData)
            : letterheadData;
          const result = await mammoth.convertToHtml(
            { arrayBuffer },
            { convertImage: mammoth.images.imgElement(image => image.read('base64').then(data => ({
              src: 'data:' + image.contentType + ';base64,' + data,
            }))) },
          );
          if (cancelled) return;
          const html = (result.value || '').trim();
          const hasText = html.replace(/<[^>]*>/g, '').trim().length > 0;
          const hasImage = /<img\b/i.test(html);
          if (!hasText && !hasImage) {
            setError('The Word file opened, but no visible text or images could be extracted. Re-save it as .docx in Microsoft Word and upload it again.');
          } else {
            setFallbackHtml(html);
            if (container) container.replaceChildren();
          }
          setLoading(false);
        } catch (fallbackError) {
          if (cancelled) return;
          console.error('Both DOCX preview methods failed:', fallbackError);
          setError('Unable to render this Word document in the browser. You can still download it and open it in Microsoft Word.');
          setLoading(false);
        }
      } finally {
        if (timeoutId) window.clearTimeout(timeoutId);
      }
    }

    renderDocument();
    return () => {
      cancelled = true;
      if (timeoutId) window.clearTimeout(timeoutId);
    };
  }, [letterheadData]);

  return (
    <div className="letterhead-preview-component">
      <div className="uploaded-file-label">
        <span>Original Word letterhead preview{fileName ? ' · ' + fileName : ''}</span>
        {loading && <span className="preview-rendering">Rendering…</span>}
      </div>
      {loading && <p className="preview-status">Loading original document formatting, pages, tables and images…</p>}
      {error && <p className="preview-error">{error}</p>}
      <div ref={previewContainerRef} className={fallbackHtml ? 'docx-rendered hidden' : 'docx-rendered'} />
      {fallbackHtml && <div className="docx-rendered docx-fallback" dangerouslySetInnerHTML={{ __html: fallbackHtml }} />}
      <p className="preview-disclaimer">This is a browser preview. Check the downloaded file in Microsoft Word for exact pagination and layout rendering.</p>
    </div>
  );
}
