import type {
  PptxViewer,
  SlideHandle,
  TextSearchResult
} from '@aiden0z/pptx-renderer';
import { ISignal, Signal } from '@lumino/signaling';
import { IDocumentViewer, ZOOM_STEP, formatFindStatus } from './viewer';

const THUMBNAIL_WIDTH = 120;

/**
 * A PPTX deck shown one slide at a time by `@aiden0z/pptx-renderer`. Every
 * operation here is a call into the viewer: navigation, zoom, thumbnails and
 * search are its features, this class only wires them to the widget.
 */
export class SlideDeck implements IDocumentViewer {
  /**
   * Parse `bytes` and render the first slide into `host`
   */
  static async open(bytes: Uint8Array, host: HTMLElement): Promise<SlideDeck> {
    const { PptxViewer, RECOMMENDED_ZIP_LIMITS } = await import(
      '@aiden0z/pptx-renderer'
    );
    const thumbnails = document.createElement('div');
    thumbnails.className = 'jp-DocReaderWidget-thumbnails';
    const stage = document.createElement('div');
    stage.className = 'jp-DocReaderWidget-stage';
    host.append(thumbnails, stage);
    const viewer = await PptxViewer.open(bytes, stage, {
      renderMode: 'slide',
      zipLimits: RECOMMENDED_ZIP_LIMITS,
      pdfjs: false
    });
    return new SlideDeck(viewer, thumbnails);
  }

  private constructor(viewer: PptxViewer, thumbnails: HTMLElement) {
    this._viewer = viewer;
    this._thumbnails = thumbnails;
    this._observer = new IntersectionObserver(
      entries => this._renderThumbnails(entries),
      { root: thumbnails }
    );
    for (let i = 0; i < viewer.slideCount; i++) {
      const item = document.createElement('button');
      item.className = 'jp-DocReaderWidget-thumbnail';
      item.title = `Slide ${i + 1}`;
      item.dataset.index = String(i);
      item.addEventListener('click', () => void this.goTo(i));
      thumbnails.append(item);
      this._observer.observe(item);
    }
    viewer.on('slidechange', () => this._onSlideChange());
    this._onSlideChange();
  }

  /**
   * Emitted when the shown slide or the search status changes
   */
  get changed(): ISignal<this, void> {
    return this._changed;
  }

  /**
   * Index of the shown slide, 0-based
   */
  get index(): number {
    return this._viewer.currentSlideIndex;
  }

  /**
   * Number of slides in the deck
   */
  get count(): number {
    return this._viewer.slideCount;
  }

  /**
   * Result of the last search: empty before any, else `k of m` or `No matches`
   */
  get findStatus(): string {
    return formatFindStatus(this._query, this._match, this._matches.length);
  }

  /**
   * Show slide `index`, clamped to the deck by the viewer
   */
  async goTo(index: number): Promise<void> {
    await this._viewer.goToSlide(index);
  }

  /**
   * Change the zoom by `steps` of 25 percent of the fitted width
   */
  async zoom(steps: number): Promise<void> {
    await this._viewer.setZoom(
      this._viewer.zoomPercent + steps * ZOOM_STEP * 100
    );
  }

  /**
   * Fit the slide to the width of the panel, the zoom it opened at
   */
  async resetZoom(): Promise<void> {
    await this._viewer.setZoom(100);
  }

  /**
   * Show and highlight the next match of `query`, or the previous one when
   * `backwards` is set. A new query starts from the first match.
   */
  async find(query: string, backwards = false): Promise<void> {
    if (query !== this._query) {
      this._query = query;
      this._matches = query === '' ? [] : this._viewer.searchText(query);
      this._match = -1;
    }
    this._viewer.clearSearchHighlights();
    const count = this._matches.length;
    if (count > 0) {
      this._match = backwards
        ? (Math.max(this._match, 0) - 1 + count) % count
        : (this._match + 1) % count;
      const result = this._matches[this._match];
      await this._viewer.goToSlide(result.slideIndex);
      await this._viewer.highlightSearchResult(result, {
        className: 'jp-DocReaderWidget-match'
      });
    }
    this._changed.emit();
  }

  /**
   * Release the viewer, its thumbnails and observers
   */
  dispose(): void {
    this._observer.disconnect();
    this._thumbnailHandles.forEach(handle => handle.dispose());
    this._viewer.destroy();
    Signal.clearData(this);
  }

  private _renderThumbnails(entries: IntersectionObserverEntry[]): void {
    for (const entry of entries) {
      if (!entry.isIntersecting) {
        continue;
      }
      const item = entry.target as HTMLElement;
      this._observer.unobserve(item);
      const handle = this._viewer.renderThumbnailToContainer(
        Number(item.dataset.index),
        item,
        { width: THUMBNAIL_WIDTH }
      );
      if (handle) {
        this._thumbnailHandles.push(handle);
      }
    }
  }

  private _onSlideChange(): void {
    this._thumbnails.childNodes.forEach((item, i) => {
      const active = i === this.index;
      (item as HTMLElement).classList.toggle('jp-mod-active', active);
      (item as HTMLElement).setAttribute('aria-current', String(active));
      if (active) {
        (item as HTMLElement).scrollIntoView({ block: 'nearest' });
      }
    });
    this._changed.emit();
  }

  private _viewer: PptxViewer;
  private _thumbnails: HTMLElement;
  private _observer: IntersectionObserver;
  private _thumbnailHandles: SlideHandle[] = [];
  private _query = '';
  private _matches: TextSearchResult[] = [];
  private _match = -1;
  private _changed = new Signal<this, void>(this);
}
