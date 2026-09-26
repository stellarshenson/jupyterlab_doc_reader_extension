import type { Odr } from '@opendocument/odr-core';
import { ISignal, Signal } from '@lumino/signaling';
import { EXTERNAL_LINK } from './render';
import { IDocumentViewer, ZOOM_STEP, formatFindStatus } from './viewer';

/**
 * Runs inside the sandboxed frame. The frame has an opaque origin, so the
 * host reaches the viewer's own search and zoom functions (`window.odr`) only
 * through these messages; link clicks go to the host, which opens web and
 * mail links and drops the rest.
 */
const BRIDGE = `<script>
(function () {
  window.addEventListener('message', function (event) {
    if (event.source !== window.parent) {
      return;
    }
    var odr = window.odr;
    var request = event.data;
    var result = 0;
    if (request.command === 'find') {
      if (request.fresh) {
        result = odr.search(request.query);
      } else {
        result = request.backwards ? odr.searchPrevious() : odr.searchNext();
      }
    } else if (request.command === 'zoom') {
      result = odr.setZoom(odr.getZoom() + request.step);
    } else if (request.command === 'fit') {
      result = odr.resetZoom();
    }
    window.parent.postMessage({ id: request.id, result: result }, '*');
  });
  document.addEventListener('click', function (event) {
    var link = event.target.closest && event.target.closest('a[href]');
    if (link && link.getAttribute('href').charAt(0) !== '#') {
      event.preventDefault();
      window.parent.postMessage({ link: link.getAttribute('href') }, '*');
    }
  }, true);
})();
</script>`;

let odrModule: Promise<Odr> | null = null;

/**
 * Convert ODT or ODP bytes to one self-contained HTML page with
 * `@opendocument/odr-core`. The wasm module is loaded once and kept.
 */
async function renderOdfPage(
  bytes: Uint8Array,
  format: 'odt' | 'odp'
): Promise<string> {
  if (!odrModule) {
    odrModule = import('@opendocument/odr-core').then(({ Odr }) => Odr.load());
  }
  const odr = await odrModule;
  // The type is forced: detection by content would show any other file,
  // plain text included, instead of reporting that it is not ODF
  const document = odr.open(bytes, {
    fileType: odr.enums.FileType[format],
    // the viewer's own page frame for text, instead of one bare column
    textDocumentMargin: true
  });
  try {
    return document.render(0).html;
  } finally {
    document.close();
  }
}

/**
 * An ODT or ODP document shown by `@opendocument/odr-core` in a frame
 * sandboxed with `allow-scripts` only: the page and its script run in an
 * opaque origin and cannot reach or navigate JupyterLab.
 */
export class OdfFrame implements IDocumentViewer {
  /**
   * Render `bytes` into a sandboxed frame in `host`
   */
  static async open(
    bytes: Uint8Array,
    format: 'odt' | 'odp',
    host: HTMLElement
  ): Promise<OdfFrame> {
    const page = await renderOdfPage(bytes, format);
    const frame = document.createElement('iframe');
    frame.className = 'jp-DocReaderWidget-frame';
    frame.title =
      format === 'odp' ? 'OpenDocument presentation' : 'OpenDocument text';
    frame.setAttribute('sandbox', 'allow-scripts');
    frame.srcdoc = page + BRIDGE;
    const loaded = new Promise(resolve =>
      frame.addEventListener('load', resolve, { once: true })
    );
    host.append(frame);
    await loaded;
    return new OdfFrame(frame);
  }

  private constructor(frame: HTMLIFrameElement) {
    this._frame = frame;
    window.addEventListener('message', this);
  }

  get changed(): ISignal<this, void> {
    return this._changed;
  }

  get findStatus(): string {
    return formatFindStatus(this._query, this._match, this._count);
  }

  /**
   * Handle the messages of this frame's page
   */
  handleEvent(event: MessageEvent): void {
    if (event.source !== this._frame.contentWindow) {
      return;
    }
    const { id, result, link } = event.data ?? {};
    if (typeof link === 'string') {
      if (EXTERNAL_LINK.test(link.trim())) {
        window.open(link.trim(), '_blank', 'noopener,noreferrer');
      }
      return;
    }
    this._pending.get(id)?.(Number(result));
    this._pending.delete(id);
  }

  /**
   * Show the next match of `query` with the viewer's search, or the previous
   * one when `backwards` is set. A new query runs the viewer's search, which
   * selects the first match; the same query steps the selection.
   */
  async find(query: string, backwards: boolean): Promise<void> {
    const fresh = query !== this._query;
    // recorded before the await, so the next Enter compares against the
    // query last sent to the page, not the last one answered
    this._query = query;
    const count = await this._call({
      command: 'find',
      query,
      backwards,
      fresh
    });
    if (fresh) {
      this._match = 0;
    } else if (count > 0) {
      this._match = (this._match + (backwards ? count - 1 : 1)) % count;
    }
    this._count = count;
    this._changed.emit();
  }

  async zoom(steps: number): Promise<void> {
    await this._call({ command: 'zoom', step: steps * ZOOM_STEP });
  }

  async resetZoom(): Promise<void> {
    await this._call({ command: 'fit' });
  }

  dispose(): void {
    window.removeEventListener('message', this);
    this._pending.clear();
    Signal.clearData(this);
  }

  /**
   * Send a command to the page and wait for its result
   */
  private _call(request: Record<string, unknown>): Promise<number> {
    const id = ++this._lastId;
    return new Promise(resolve => {
      this._pending.set(id, resolve);
      this._frame.contentWindow?.postMessage({ ...request, id }, '*');
    });
  }

  private _frame: HTMLIFrameElement;
  private _pending = new Map<number, (result: number) => void>();
  private _lastId = 0;
  private _query = '';
  private _count = 0;
  private _match = 0;
  private _changed = new Signal<this, void>(this);
}
