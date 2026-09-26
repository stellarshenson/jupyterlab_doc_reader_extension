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
import { OdfFrame } from './odf';
import { decodeBase64, guardLinks, renderDocx, renderRtf } from './render';
import { SlideDeck } from './slides';
import { IDocumentViewer } from './viewer';
import { Workbook } from './workbook';

/**
 * Binary formats no browser renderer reads, and the format to save them as
 */
const LEGACY_FORMATS: Record<string, string> = {
  '.doc': 'DOCX',
  '.ppt': 'PPTX'
};

/**
 * Formats whose viewer offers zoom and search, with their toolbar text: the
 * button that returns to the opening zoom, and the find box
 */
const TOOLBAR_TEXT: Record<
  string,
  { reset: string; resetTooltip: string; find: string }
> = {
  '.pptx': {
    reset: 'Fit',
    resetTooltip: 'Fit slide to width',
    find: 'Find in slides'
  },
  '.odp': {
    reset: 'Fit',
    resetTooltip: 'Fit slide to width, never above 100%',
    find: 'Find in slides'
  },
  '.odt': {
    reset: 'Fit',
    resetTooltip: 'Fit page to width, never above 100%',
    find: 'Find in document'
  },
  '.xlsx': {
    reset: '100%',
    resetTooltip: 'Zoom to 100%',
    find: 'Find in sheets'
  }
};

/**
 * A widget that renders DOCX, PPTX, RTF, ODT, ODP and XLSX files in the
 * browser, from the bytes JupyterLab loads for the document
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
   * The viewer of a PPTX, ODT, ODP or XLSX document, once rendered
   */
  get viewer(): IDocumentViewer | null {
    return this._viewer;
  }

  /**
   * Emitted when the document is rendered and whenever its viewer changes
   */
  get viewerChanged(): ISignal<this, void> {
    return this._viewerChanged;
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
    this._viewer?.dispose();
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
      // slide from the width of its container, and a frame loads its page
      // only once it is in the document
      const host = document.createElement('div');
      host.className = `jp-DocReaderWidget-content jp-DocReaderWidget-${ext.slice(1)}`;
      this.node.append(host);
      let viewer: IDocumentViewer | null = null;
      if (ext === '.pptx') {
        viewer = await SlideDeck.open(bytes, host);
      } else if (ext === '.odt' || ext === '.odp') {
        viewer = await OdfFrame.open(
          bytes,
          ext === '.odt' ? 'odt' : 'odp',
          host
        );
      } else if (ext === '.xlsx') {
        viewer = await Workbook.open(bytes, host);
      } else if (ext === '.rtf') {
        await renderRtf(bytes, host);
        guardLinks(host);
      } else {
        await renderDocx(bytes, host);
        guardLinks(host);
      }
      // the tab may have been closed while the viewer read the file
      if (this.isDisposed) {
        viewer?.dispose();
        return;
      }
      if (viewer) {
        this._viewer = viewer;
        viewer.changed.connect(() => this._viewerChanged.emit());
      }
      // Remove only the loading message: moving the host would detach it,
      // and a detached frame reloads its page
      Array.from(this.node.children)
        .filter(child => child !== host)
        .forEach(child => child.remove());
      this._viewerChanged.emit();
    } catch (error) {
      this._viewer?.dispose();
      this._viewer = null;
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
    const deck = this._viewer;
    if (!(deck instanceof SlideDeck)) {
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
  private _viewer: IDocumentViewer | null = null;
  private _viewerChanged = new Signal<this, void>(this);
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
    const ext = PathExt.extname(context.path).toLowerCase();
    if (TOOLBAR_TEXT[ext]) {
      addViewerToolbar(widget.toolbar, content, ext);
    }
    return widget;
  }
}

/**
 * Add the viewer controls: previous, position and next for slides, then zoom,
 * the opening zoom and find, each a call into the format's viewer
 */
function addViewerToolbar(
  toolbar: Toolbar,
  content: DocReaderWidget,
  ext: string
): void {
  const text = TOOLBAR_TEXT[ext];
  const button = (
    name: string,
    options: ToolbarButtonComponent.IProps,
    action: (viewer: IDocumentViewer) => Promise<void>
  ) => {
    toolbar.addItem(
      name,
      new ToolbarButton({
        ...options,
        // keep focus on the slide so the navigation keys keep working
        noFocusOnClick: true,
        onClick: () => {
          if (content.viewer) {
            void action(content.viewer);
          }
        }
      })
    );
  };
  const position = new Widget({ node: document.createElement('span') });
  if (ext === '.pptx') {
    const move = (step: number) => async (viewer: IDocumentViewer) => {
      if (viewer instanceof SlideDeck) {
        await viewer.goTo(viewer.index + step);
      }
    };
    button(
      'previous-slide',
      { icon: caretLeftIcon, tooltip: 'Previous slide' },
      move(-1)
    );
    position.addClass('jp-DocReaderWidget-position');
    position.node.setAttribute('role', 'status');
    toolbar.addItem('slide-position', position);
    button(
      'next-slide',
      { icon: caretRightIcon, tooltip: 'Next slide' },
      move(1)
    );
  }
  button('zoom-out', { label: '-', tooltip: 'Zoom out' }, viewer =>
    viewer.zoom(-1)
  );
  button('zoom-in', { label: '+', tooltip: 'Zoom in' }, viewer =>
    viewer.zoom(1)
  );
  button(
    'reset-zoom',
    { label: text.reset, tooltip: text.resetTooltip },
    viewer => viewer.resetZoom()
  );
  toolbar.addItem('spacer', Toolbar.createSpacerItem());

  const find = new Widget();
  find.addClass('jp-DocReaderWidget-find');
  const input = document.createElement('input');
  input.type = 'search';
  input.placeholder = text.find;
  input.setAttribute('aria-label', text.find);
  input.title = 'Enter for the next match, Shift+Enter for the previous one';
  const status = document.createElement('span');
  status.className = 'jp-DocReaderWidget-findStatus';
  status.setAttribute('role', 'status');
  find.node.append(input, status);
  input.addEventListener('keydown', event => {
    if (event.key === 'Enter' && content.viewer) {
      event.preventDefault();
      void content.viewer.find(input.value, event.shiftKey);
    }
  });
  toolbar.addItem('find', find);

  content.viewerChanged.connect(() => {
    const viewer = content.viewer;
    if (viewer instanceof SlideDeck) {
      position.node.textContent = `${viewer.index + 1} / ${viewer.count}`;
    }
    status.textContent = viewer?.findStatus ?? '';
  });
}
