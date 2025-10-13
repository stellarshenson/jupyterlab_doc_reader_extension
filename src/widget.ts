import { DocumentWidget } from '@jupyterlab/docregistry';
import { ABCWidgetFactory, DocumentRegistry } from '@jupyterlab/docregistry';
import { PromiseDelegate } from '@lumino/coreutils';
import { Widget } from '@lumino/widgets';
import { requestAPI } from './handler';

/**
 * A widget for displaying document files (DOCX, DOC, RTF) as PDFs
 */
export class DocReaderWidget extends Widget {
  private _context: DocumentRegistry.Context;
  private _ready = new PromiseDelegate<void>();
  private _iframe: HTMLIFrameElement;
  private _errorDiv: HTMLDivElement;

  constructor(context: DocumentRegistry.Context) {
    super();
    this._context = context;

    this.addClass('jp-DocReaderWidget');
    this.title.label = context.localPath;

    // Create iframe for PDF display
    this._iframe = document.createElement('iframe');
    this._iframe.className = 'jp-DocReaderWidget-iframe';
    this._iframe.style.width = '100%';
    this._iframe.style.height = '100%';
    this._iframe.style.border = 'none';

    // Create error display div
    this._errorDiv = document.createElement('div');
    this._errorDiv.className = 'jp-DocReaderWidget-error';
    this._errorDiv.style.display = 'none';
    this._errorDiv.style.padding = '20px';
    this._errorDiv.style.color = '#d32f2f';

    this.node.appendChild(this._errorDiv);
    this.node.appendChild(this._iframe);

    // Load and convert the document
    void this._loadDocument();
  }

  /**
   * A promise that resolves when the widget is ready.
   */
  get ready(): Promise<void> {
    return this._ready.promise;
  }

  /**
   * Load the document and convert it to PDF
   */
  private async _loadDocument(): Promise<void> {
    try {
      const path = this._context.path;

      // Show loading indicator
      this._iframe.srcdoc = `
        <html>
          <body style="display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; font-family: sans-serif;">
            <div style="text-align: center;">
              <div style="font-size: 18px; margin-bottom: 10px;">Converting document...</div>
              <div style="font-size: 14px; color: #666;">Please wait while we process your file</div>
            </div>
          </body>
        </html>
      `;

      // Request conversion from server
      console.log('[DocReader] Requesting conversion for:', path);
      let response: any;
      try {
        response = await requestAPI<any>('convert', {
          method: 'POST',
          body: JSON.stringify({ path })
        });
        console.log('[DocReader] Received response:', {
          success: response?.success,
          hasPdfData: !!response?.pdf_data,
          pdfDataLength: response?.pdf_data?.length,
          filename: response?.filename
        });
      } catch (apiError: any) {
        // Try to parse error response
        console.error('[DocReader] API request failed:', apiError);
        if (apiError?.message) {
          try {
            const errorData = JSON.parse(apiError.message);
            throw errorData;
          } catch {
            // If parsing fails, throw the original error
            throw apiError;
          }
        }
        throw apiError;
      }

      if (response.success && response.pdf_data) {
        console.log('[DocReader] Processing PDF data...');
        // Create a blob from the base64 PDF data
        console.log('[DocReader] Converting base64 to blob...');
        const pdfBlob = this._base64ToBlob(
          response.pdf_data,
          'application/pdf'
        );
        console.log('[DocReader] Created blob. Size:', pdfBlob.size, 'bytes');

        const pdfUrl = URL.createObjectURL(pdfBlob);
        console.log('[DocReader] Created object URL:', pdfUrl);

        // Display the PDF using embed tag (better compatibility than iframe)
        // Clear previous content
        this._iframe.srcdoc = '';
        this._iframe.style.display = 'none';

        // Create embed element for PDF
        const embed = document.createElement('embed');
        embed.src = pdfUrl;
        embed.type = 'application/pdf';
        embed.style.width = '100%';
        embed.style.height = '100%';
        embed.style.border = 'none';

        // Clear the widget and add embed
        while (this.node.firstChild) {
          this.node.removeChild(this.node.firstChild);
        }
        this.node.appendChild(embed);

        console.log('[DocReader] Set embed src to blob URL');

        this._ready.resolve();
        console.log('[DocReader] Document loading complete');
      } else {
        // Response indicates failure
        console.error('[DocReader] Response missing success or pdf_data:', response);
        throw response;
      }
    } catch (error) {
      console.error('Error loading document:', error);
      this._showError(error);
      this._ready.reject(error);
    }
  }

  /**
   * Convert base64 string to Blob
   */
  private _base64ToBlob(base64: string, contentType: string): Blob {
    const byteCharacters = atob(base64);
    const byteNumbers = new Array(byteCharacters.length);

    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }

    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: contentType });
  }

  /**
   * Show error message
   */
  private _showError(error: any): void {
    this._iframe.style.display = 'none';
    this._errorDiv.style.display = 'block';

    console.error('Document conversion error:', error);

    // Try to extract detailed error information
    let errorMessage = 'Unknown error occurred';
    let errorType = 'Error';
    let traceback = '';
    let debugInfo = '';

    if (error?.error) {
      errorMessage = error.error;
    } else if (error?.message) {
      errorMessage = error.message;
    } else if (typeof error === 'string') {
      errorMessage = error;
    } else {
      errorMessage = JSON.stringify(error);
    }

    if (error?.error_type) {
      errorType = error.error_type;
    }

    if (error?.traceback) {
      traceback = error.traceback;
    }

    if (error?.file_path || error?.full_path) {
      debugInfo = `
        <div style="margin-top: 15px; padding: 10px; background: #f5f5f5; border-radius: 4px;">
          <strong>Debug Information:</strong><br>
          ${error.file_path ? `<div style="margin: 5px 0;"><strong>Requested path:</strong> <code>${error.file_path}</code></div>` : ''}
          ${error.full_path ? `<div style="margin: 5px 0;"><strong>Full path:</strong> <code>${error.full_path}</code></div>` : ''}
        </div>
      `;
    }

    const tracebackSection = traceback
      ? `
        <details style="margin-top: 15px;">
          <summary style="cursor: pointer; font-weight: bold; padding: 10px; background: #f5f5f5; border-radius: 4px;">
            View Full Traceback
          </summary>
          <pre style="margin: 10px 0; padding: 10px; background: #2d2d2d; color: #f8f8f8; border-radius: 4px; overflow-x: auto; font-size: 12px; line-height: 1.4;">${this._escapeHtml(
            traceback
          )}</pre>
        </details>
      `
      : '';

    this._errorDiv.innerHTML = `
      <div style="padding: 20px; font-family: sans-serif; overflow-y: auto; max-height: 100%;">
        <h3 style="color: #d32f2f; margin: 0 0 15px 0;">Failed to load document</h3>

        <div style="margin-bottom: 15px;">
          <strong>Error Type:</strong> <code>${this._escapeHtml(errorType)}</code>
        </div>

        <div style="margin-bottom: 15px; padding: 10px; background: #ffebee; border-left: 4px solid #d32f2f; border-radius: 4px;">
          <strong>Error Message:</strong><br>
          <pre style="margin: 5px 0; white-space: pre-wrap; font-family: monospace;">${this._escapeHtml(
            errorMessage
          )}</pre>
        </div>

        ${debugInfo}
        ${tracebackSection}

        <div style="margin-top: 20px; padding: 15px; background: #e3f2fd; border-radius: 4px;">
          <strong>Troubleshooting:</strong>
          <ul style="margin: 10px 0; padding-left: 20px;">
            <li>Check that mammoth and weasyprint are installed: <code>pip list | grep -E "mammoth|weasyprint"</code></li>
            <li>Verify the extension is properly installed: <code>jupyter server extension list</code></li>
            <li>Check the file is a valid DOCX, DOC, or RTF document</li>
            <li>Ensure the file is not corrupted or password-protected</li>
            <li>For legacy DOC files, consider converting to DOCX format</li>
            <li>Check JupyterLab logs for more details: <code>/var/log/jupyterlab.log</code></li>
          </ul>
        </div>
      </div>
    `;
  }

  /**
   * Escape HTML to prevent XSS
   */
  private _escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * Dispose of the resources held by the widget.
   */
  dispose(): void {
    if (this.isDisposed) {
      return;
    }

    // Clean up the blob URL if it exists
    if (this._iframe.src.startsWith('blob:')) {
      URL.revokeObjectURL(this._iframe.src);
    }

    super.dispose();
  }
}

/**
 * A widget factory for document readers.
 */
export class DocReaderFactory extends ABCWidgetFactory<
  DocumentWidget<DocReaderWidget>,
  DocumentRegistry.IModel
> {
  /**
   * Create a new widget given a context.
   */
  protected createNewWidget(
    context: DocumentRegistry.Context
  ): DocumentWidget<DocReaderWidget> {
    const content = new DocReaderWidget(context);
    const widget = new DocumentWidget({ content, context });

    return widget;
  }
}
