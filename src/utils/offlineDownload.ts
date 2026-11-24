import { shareOrDownloadBlob } from './platform';

export async function downloadOfflinePackage(target: 'windows' | 'ipad'): Promise<string> {
  const endpoint = target === 'ipad' ? '/offline-ipad.dat' : '/offline-windows.dat';
  const res = await fetch(endpoint, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }

  if (target === 'ipad') {
    const buf = await res.arrayBuffer();
    const bytes = new Uint8Array(buf);
    // Verify ZIP magic header ("PK\x03\x04")
    if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      throw new Error('iPad offline zip package unavailable');
    }
    const filename = 'Grand_Group_iPad_App.zip';
    const blob = new Blob([buf], { type: 'application/zip' });
    await shareOrDownloadBlob(blob, filename, 'application/zip');
    return filename;
  }

  const htmlText = await res.text();
  if (!htmlText.trim().toLowerCase().startsWith('<!doctype html')) {
    throw new Error('Windows offline package unavailable');
  }
  const filename = 'Grand_Group_FZC_Windows_Offline.html';
  const blob = new Blob([htmlText], { type: 'text/html;charset=utf-8' });
  await shareOrDownloadBlob(blob, filename, 'text/html');
  return filename;
}

export async function downloadTutorialAsset(
  asset: 'pdf' | 'video' | 'markdown'
): Promise<string> {
  const map = {
    pdf: {
      endpoint: '/notebooklm-tutorial-pdf.dat',
      filename: 'Grand_Group_Procurement_NotebookLM_Tutorial_Guide.pdf',
      mime: 'application/pdf',
    },
    video: {
      endpoint: '/demo-walkthrough.dat',
      filename: 'Grand_Group_Procurement_Demo_Walkthrough.webm',
      mime: 'video/webm',
    },
    markdown: {
      endpoint: '/notebooklm-tutorial-md.dat',
      filename: 'Grand_Group_Procurement_NotebookLM_Guide.md',
      mime: 'text/markdown;charset=utf-8',
    },
  }[asset];

  const res = await fetch(map.endpoint, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  const buf = await res.arrayBuffer();
  const blob = new Blob([buf], { type: map.mime });
  await shareOrDownloadBlob(blob, map.filename, map.mime);
  return map.filename;
}
