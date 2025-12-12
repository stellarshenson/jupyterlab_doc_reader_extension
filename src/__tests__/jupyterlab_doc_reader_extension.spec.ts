/**
 * Unit tests for jupyterlab_doc_reader_extension
 * Note: JupyterLab extension integration tests require Playwright (see ui-tests/)
 */

describe('jupyterlab_doc_reader_extension', () => {
  it('should be tested', () => {
    expect(1 + 1).toEqual(2);
  });
});

describe('base64 conversion', () => {
  it('should correctly decode base64 to bytes', () => {
    // Test the base64 to blob conversion logic
    const testBase64 = btoa('Hello, World!');
    const byteCharacters = atob(testBase64);
    const byteNumbers = new Array(byteCharacters.length);

    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }

    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: 'text/plain' });

    expect(blob.size).toBe(13); // "Hello, World!" is 13 bytes
  });

  it('should create blob with correct type', () => {
    const testBase64 = btoa('test');
    const byteCharacters = atob(testBase64);
    const byteNumbers = new Array(byteCharacters.length);

    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }

    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: 'application/pdf' });

    expect(blob.type).toBe('application/pdf');
  });
});

describe('HTML escaping', () => {
  it('should escape HTML special characters', () => {
    const escapeHtml = (text: string): string => {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    };

    expect(escapeHtml('<script>alert("xss")</script>')).toBe(
      '&lt;script&gt;alert("xss")&lt;/script&gt;'
    );
    expect(escapeHtml('a & b')).toBe('a &amp; b');
    expect(escapeHtml('"quoted"')).toBe('"quoted"');
  });
});

describe('file extension detection', () => {
  it('should identify PPTX files', () => {
    const path = '/path/to/presentation.pptx';
    const ext = path.split('.').pop()?.toLowerCase();
    expect(ext).toBe('pptx');
  });

  it('should identify DOCX files', () => {
    const path = '/path/to/document.docx';
    const ext = path.split('.').pop()?.toLowerCase();
    expect(ext).toBe('docx');
  });

  it('should identify RTF files', () => {
    const path = '/path/to/document.rtf';
    const ext = path.split('.').pop()?.toLowerCase();
    expect(ext).toBe('rtf');
  });

  it('should handle uppercase extensions', () => {
    const path = '/path/to/PRESENTATION.PPTX';
    const ext = path.split('.').pop()?.toLowerCase();
    expect(ext).toBe('pptx');
  });
});
