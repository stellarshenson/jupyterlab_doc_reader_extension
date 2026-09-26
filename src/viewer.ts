import { ISignal } from '@lumino/signaling';

/**
 * A rendered document whose viewer library offers zoom and search. The
 * toolbar calls these; the extension adds no zoom or search of its own.
 */
export interface IDocumentViewer {
  /**
   * Emitted when the shown position or the search status changes
   */
  readonly changed: ISignal<IDocumentViewer, void>;

  /**
   * Result of the last search: empty before any, else `k of m` or `No matches`
   */
  readonly findStatus: string;

  /**
   * Show the next match of `query`, or the previous one when `backwards` is set
   */
  find(query: string, backwards: boolean): Promise<void>;

  /**
   * Change the zoom by `steps` of 25 percent
   */
  zoom(steps: number): Promise<void>;

  /**
   * Return to the zoom the viewer opened at
   */
  resetZoom(): Promise<void>;

  dispose(): void;
}

/**
 * The zoom change of one toolbar step
 */
export const ZOOM_STEP = 0.25;

/**
 * Format the search status every viewer shows in the toolbar
 */
export function formatFindStatus(
  query: string,
  index: number,
  count: number
): string {
  if (query === '') {
    return '';
  }
  return count === 0 ? 'No matches' : `${index + 1} of ${count}`;
}
