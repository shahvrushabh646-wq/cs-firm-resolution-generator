// Convert a File or Blob to a data URL suitable for localStorage persistence.
export const fileToBase64 = (file) => new Promise((resolve, reject) => {
  if (!file) {
    reject(new Error('No file was selected.'));
    return;
  }
  const reader = new FileReader();
  reader.onload = () => typeof reader.result === 'string'
    ? resolve(reader.result)
    : reject(new Error('Could not encode the selected file.'));
  reader.onerror = () => reject(reader.error || new Error('Could not read the selected file.'));
  reader.readAsDataURL(file);
});

// Accepts either a data URL or a raw Base64 string.
export const base64ToArrayBuffer = (base64) => {
  const raw = String(base64 || '').trim();
  const payload = raw.includes(',') ? raw.slice(raw.indexOf(',') + 1) : raw;
  if (!payload) throw new Error('The saved Word file is empty. Please upload the .docx again.');
  const binaryString = window.atob(payload.replace(/\s/g, ''));
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
};

/**
 * Safely triggers a browser download from a Blob.
 */
export const downloadBlob = (blob, filename = 'Board_Resolution.docx') => {
  if (!(blob instanceof Blob)) {
    throw new Error('The generated document is not a valid file.');
  }
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename || 'Board_Resolution.docx';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
};
