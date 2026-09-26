/**
 * One-shot renderers for DOCX and RTF, plus the link guard their output goes
 * through. PPTX links are left to the PPTX viewer.
 */

let nextDocxId = 0;

/**
 * Decode the base64 content JupyterLab loads for a binary file
 */
export function decodeBase64(data: string): Uint8Array {
  const binary = atob(data);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * Render a DOCX file into `host` with docx-preview
 */
export async function renderDocx(
  bytes: Uint8Array,
  host: HTMLElement
): Promise<void> {
  const { renderAsync } = await import('docx-preview');
  await renderAsync(bytes, host, undefined, {
    // docx-preview writes global CSS keyed by this prefix; a prefix per
    // document keeps two open documents from restyling each other
    className: `jp-DocReaderDocx${nextDocxId++}`,
    // data URLs instead of blob URLs, so nothing is left to revoke
    useBase64URL: true,
    // altChunk parts are HTML the library puts in an unsandboxed iframe,
    // where it would run script in the JupyterLab origin
    renderAltChunks: false
  });
}

/**
 * Render an RTF file into `host` with rtf.js
 */
export async function renderRtf(
  bytes: Uint8Array,
  host: HTMLElement
): Promise<void> {
  const { RTFJS, WMFJS, EMFJS } = await import('rtf.js');
  RTFJS.loggingEnabled(false);
  WMFJS.loggingEnabled(false);
  EMFJS.loggingEnabled(false);
  const document = new RTFJS.Document(bytes.buffer as ArrayBuffer, {});
  host.append(...(await document.render()));
}

export const EXTERNAL_LINK = /^(https?:|mailto:)/i;

/**
 * Stop document links from navigating JupyterLab itself: external links open
 * in a new tab, `#` links scroll inside the document, any other link is inert
 */
export function guardLinks(host: HTMLElement): void {
  host.querySelectorAll('a').forEach(link => {
    const href = (link.getAttribute('href') ?? '').trim();
    if (EXTERNAL_LINK.test(href)) {
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
    } else if (href.length < 2 || !href.startsWith('#')) {
      link.removeAttribute('href');
      link.title = 'Link disabled: only web, mail and in-document links open';
    }
  });
  host.addEventListener('click', event => {
    const link = (event.target as Element).closest('a[href^="#"]');
    if (!link) {
      return;
    }
    event.preventDefault();
    const id = link.getAttribute('href')!.slice(1);
    Array.from(host.querySelectorAll('[id]'))
      .find(element => element.id === id)
      ?.scrollIntoView();
  });
}
