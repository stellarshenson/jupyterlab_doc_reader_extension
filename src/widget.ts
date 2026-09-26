import { PathExt } from '@jupyterlab/coreutils';
import {
  ABCWidgetFactory,
  DocumentRegistry,
  DocumentWidget
} from '@jupyterlab/docregistry';
import {
  Toolbar,
  ToolbarButton,
  ToolbarButtonComponent,
  caretLeftIcon,
  caretRightIcon
} from '@jupyterlab/ui-components';
import { ISignal, Signal } from '@lumino/signaling';
import { Message } from '@lumino/messaging';
import { Widget } from '@lumino/widgets';
import { decodeBase64, guardLinks, renderDocx, renderRtf } from './render';
import { SlideDeck } from './slides';

/**
 * Binary formats no browser renderer reads, and the format to save them as
 */
const LEGACY_FORMATS: Record<string, string> = {
  '.doc': 'DOCX',
  '.ppt': 'PPTX'
};

/**
 * A widget that renders DOCX, PPTX and RTF files in the browser, from the
 * bytes JupyterLab loads for the document
 */
export class DocReaderWidget extends Widget {
  constructor(context: DocumentRegistry.Context) {
    super();
    this._context = context;
    this.addClass('jp-DocReaderWidget');
    this.title.label = context.localPath;
    this.node.tabIndex = -1;
    this._showMessage('jp-DocReaderWidget-loading', 'Loading document...');
    void this._load();
  }

  /**
   * The slide deck of a PPTX document, once rendered
   */
  get deck(): SlideDeck | null {
    return this._deck;
  }

  /**
   * Emitted when the deck is rendered and whenever it changes
   */
  get deckChanged(): ISignal<this, void> {
    return this._deckChanged;
  }

  /**
   * Handle the DOM events of the widget
   */
  handleEvent(event: Event): void {
    if (event.type === 'keydown') {
      this._onKeyDown(event as KeyboardEvent);
    }
  }

  dispose(): void {
    if (this.isDisposed) {
      return;
    }
    this._deck?.dispose();
    super.dispose();
  }

  protected onAfterAttach(msg: Message): void {
    super.onAfterAttach(msg);
    this.node.addEventListener('keydown', this);
  }

  protected onBeforeDetach(msg: Message): void {
    this.node.removeEventListener('keydown', this);
    super.onBeforeDetach(msg);
  }

  protected onActivateRequest(msg: Message): void {
    this.node.focus();
  }

  /**
   * Render the document, or show why it cannot be rendered
   */
  private async _load(): Promise<void> {
    try {
      await this._context.ready;
      const ext = PathExt.extname(this._context.path).toLowerCase();
      const target = LEGACY_FORMATS[ext];
      if (target) {
        this._showMessage(
          'jp-DocReaderWidget-unsupported',
          `${ext.slice(1).toUpperCase()} files are not supported`,
          `Save the file as ${target} to view it here.`
        );
        return;
      }
      const bytes = decodeBase64(this._context.model.toString());
      if (bytes.length === 0) {
        throw new Error('The file is empty.');
      }
      // The host is attached before rendering: the PPTX viewer sizes the
      // slide from the width of its container
      const host = document.createElement('div');
      host.className = `jp-DocReaderWidget-content jp-DocReaderWidget-${ext.slice(1)}`;
      this.node.append(host);
      if (ext === '.pptx') {
        const deck = await SlideDeck.open(bytes, host);
        // the tab may have been closed while the deck was parsed
        if (this.isDisposed) {
          deck.dispose();
          return;
        }
        this._deck = deck;
        this._deck.changed.connect(() => this._deckChanged.emit());
      } else if (ext === '.rtf') {
        await renderRtf(bytes, host);
        guardLinks(host);
      } else {
        await renderDocx(bytes, host);
        guardLinks(host);
      }
      this.node.replaceChildren(host);
      this._deckChanged.emit();
    } catch (error) {
      this._deck?.dispose();
      this._deck = null;
      this._showMessage(
        'jp-DocReaderWidget-error',
        'Cannot display this document',
        error instanceof Error ? error.message : String(error)
      );
    }
  }

  /**
   * Move between slides with the arrow, PageUp, PageDown, Home and End keys
   */
  private _onKeyDown(event: KeyboardEvent): void {
    const deck = this._deck;
    if (!deck) {
      return;
    }
    let index: number;
    switch (event.key) {
      case 'ArrowLeft':
      case 'PageUp':
        index = deck.index - 1;
        break;
      case 'ArrowRight':
      case 'PageDown':
        index = deck.index + 1;
        break;
      case 'Home':
        index = 0;
        break;
      case 'End':
        index = deck.count - 1;
        break;
      default:
        return;
    }
    event.preventDefault();
    void deck.goTo(index);
  }

  /**
   * Replace the content with a titled message panel
   */
  private _showMessage(className: string, title: string, detail = ''): void {
    const panel = document.createElement('div');
    panel.className = `jp-DocReaderWidget-message ${className}`;
    const heading = document.createElement('h3');
    heading.textContent = title;
    panel.append(heading);
    if (detail) {
      const text = document.createElement('p');
      text.textContent = detail;
      panel.append(text);
    }
    this.node.replaceChildren(panel);
  }

  private _context: DocumentRegistry.Context;
  private _deck: SlideDeck | null = null;
  private _deckChanged = new Signal<this, void>(this);
}

/**
 * A widget factory for document readers
 */
export class DocReaderFactory extends ABCWidgetFactory<
  DocumentWidget<DocReaderWidget>,
  DocumentRegistry.IModel
> {
  /**
   * Create a new widget given a context
   */
  protected createNewWidget(
    context: DocumentRegistry.Context
  ): DocumentWidget<DocReaderWidget> {
    const content = new DocReaderWidget(context);
    const widget = new DocumentWidget({ content, context });
    if (PathExt.extname(context.path).toLowerCase() === '.pptx') {
      addSlideToolbar(widget.toolbar, content);
    }
    return widget;
  }
}

/**
 * Add the slide viewer controls: previous, next, position, zoom, fit, find
 */
function addSlideToolbar(toolbar: Toolbar, content: DocReaderWidget): void {
  const button = (
    name: string,
    options: ToolbarButtonComponent.IProps,
    action: (deck: SlideDeck) => Promise<void>
  ) => {
    toolbar.addItem(
      name,
      new ToolbarButton({
        ...options,
        // keep focus on the slide so the navigation keys keep working
        noFocusOnClick: true,
        onClick: () => {
          if (content.deck) {
            void action(content.deck);
          }
        }
      })
    );
  };
  button(
    'previous-slide',
    { icon: caretLeftIcon, tooltip: 'Previous slide' },
    deck => deck.goTo(deck.index - 1)
  );
  const position = new Widget({ node: document.createElement('span') });
  position.addClass('jp-DocReaderWidget-position');
  position.node.setAttribute('role', 'status');
  toolbar.addItem('slide-position', position);
  button('next-slide', { icon: caretRightIcon, tooltip: 'Next slide' }, deck =>
    deck.goTo(deck.index + 1)
  );
  button('zoom-out', { label: '-', tooltip: 'Zoom out' }, deck =>
    deck.zoom(-1)
  );
  button('zoom-in', { label: '+', tooltip: 'Zoom in' }, deck => deck.zoom(1));
  button('fit', { label: 'Fit', tooltip: 'Fit slide to width' }, deck =>
    deck.fit()
  );
  toolbar.addItem('spacer', Toolbar.createSpacerItem());

  const find = new Widget();
  find.addClass('jp-DocReaderWidget-find');
  const input = document.createElement('input');
  input.type = 'search';
  input.placeholder = 'Find in slides';
  input.setAttribute('aria-label', 'Find in slides');
  input.title = 'Enter for the next match, Shift+Enter for the previous one';
  const status = document.createElement('span');
  status.className = 'jp-DocReaderWidget-findStatus';
  status.setAttribute('role', 'status');
  find.node.append(input, status);
  input.addEventListener('keydown', event => {
    if (event.key === 'Enter' && content.deck) {
      event.preventDefault();
      void content.deck.find(input.value, event.shiftKey);
    }
  });
  toolbar.addItem('find', find);

  content.deckChanged.connect(() => {
    const deck = content.deck;
    if (deck) {
      position.node.textContent = `${deck.index + 1} / ${deck.count}`;
      status.textContent = deck.findStatus;
    }
  });
}
