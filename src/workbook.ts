import type { XlsxViewer } from '@silurus/ooxml/xlsx';
import { ISignal, Signal } from '@lumino/signaling';
import { IDocumentViewer, ZOOM_STEP, formatFindStatus } from './viewer';

/**
 * Whether `bytes` hold a zip archive's closing record (end of central
 * directory): 22 bytes, followed by a comment of at most 65535 bytes
 */
function hasZipEnd(bytes: Uint8Array): boolean {
  const last = Math.max(0, bytes.length - 65557);
  for (let i = bytes.length - 22; i >= last; i--) {
    if (
      bytes[i] === 0x50 &&
      bytes[i + 1] === 0x4b &&
      bytes[i + 2] === 0x05 &&
      bytes[i + 3] === 0x06
    ) {
      return true;
    }
  }
  return false;
}

/**
 * An XLSX workbook drawn by `@silurus/ooxml` with its own sheet tabs. Zoom
 * and search are calls into that viewer.
 */
export class Workbook implements IDocumentViewer {
  /**
   * Parse `bytes` and draw the first sheet into `host`
   */
  static async open(bytes: Uint8Array, host: HTMLElement): Promise<Workbook> {
    // The viewer opens a file that is not a complete zip archive (not a zip
    // at all, or truncated) as a placeholder sheet and only logs a warning,
    // and the library exposes no parse error, so the archive's closing
    // record is checked first
    if (!hasZipEnd(bytes)) {
      throw new Error(
        'The file is not a complete zip archive, so not an XLSX workbook.'
      );
    }
    const { XlsxViewer } = await import('@silurus/ooxml/xlsx');
    // the toolbar carries the zoom, so the viewer's own slider is off
    const viewer = new XlsxViewer(host, { showZoomSlider: false });
    try {
      await viewer.load(bytes.buffer as ArrayBuffer);
    } catch (error) {
      viewer.destroy();
      throw error;
    }
    return new Workbook(viewer);
  }

  private constructor(viewer: XlsxViewer) {
    this._viewer = viewer;
  }

  get changed(): ISignal<this, void> {
    return this._changed;
  }

  get findStatus(): string {
    return formatFindStatus(this._query, this._match, this._count);
  }

  /**
   * Show the next match of `query` in any sheet, or the previous one when
   * `backwards` is set. A new query searches the workbook again.
   */
  async find(query: string, backwards: boolean): Promise<void> {
    if (query !== this._query) {
      this._query = query;
      this._viewer.clearFind();
      this._count =
        query === '' ? 0 : (await this._viewer.findText(query)).length;
    }
    if (this._count > 0) {
      const match = backwards
        ? await this._viewer.findPrev()
        : await this._viewer.findNext();
      this._match = match?.matchIndex ?? 0;
    }
    this._changed.emit();
  }

  async zoom(steps: number): Promise<void> {
    await this._viewer.setScale(this._viewer.getScale() + steps * ZOOM_STEP);
  }

  /**
   * Zoom to 100 percent, the scale the sheet opened at
   */
  async resetZoom(): Promise<void> {
    await this._viewer.setScale(1);
  }

  dispose(): void {
    this._viewer.destroy();
    Signal.clearData(this);
  }

  private _viewer: XlsxViewer;
  private _query = '';
  private _count = 0;
  private _match = 0;
  private _changed = new Signal<this, void>(this);
}
