/**
 * Unit tests for jupyterlab_doc_reader_extension.
 * Rendering itself runs in a real browser: see ui-tests/.
 */

import { decodeBase64, guardLinks } from '../render';

describe('decodeBase64', () => {
  it('decodes base64 split over lines, as Jupyter server sends it', () => {
    const text = 'PK\u0003\u0004 archive bytes';
    const encoded = btoa(text).replace(/(.{8})/g, '$1\n');
    expect(Array.from(decodeBase64(encoded))).toEqual(
      Array.from(text, c => c.charCodeAt(0))
    );
  });

  it('returns no bytes for an empty file', () => {
    expect(decodeBase64('')).toHaveLength(0);
  });
});

describe('guardLinks', () => {
  const render = (html: string): HTMLElement => {
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.append(host);
    guardLinks(host);
    return host;
  };

  afterEach(() => {
    document.body.replaceChildren();
  });

  it('opens web and mail links in a new tab', () => {
    const host = render(
      '<a href="https://jupyter.org">web</a><a href="mailto:a@b.c">mail</a>'
    );
    host.querySelectorAll('a').forEach(link => {
      expect(link.target).toBe('_blank');
      expect(link.rel).toBe('noopener noreferrer');
    });
  });

  it.each([
    ['javascript:alert(1)'],
    [' JavaScript:alert(1)'],
    ['data:text/html,x'],
    ['other.docx'],
    [''],
    ['#']
  ])('removes the link target %p', href => {
    const host = render(`<a href="${href}">link</a>`);
    const link = host.querySelector('a')!;
    expect(link.hasAttribute('href')).toBe(false);
    expect(link.title).toContain('Link disabled');
  });

  it('scrolls a # link to its bookmark inside the document', () => {
    const host = render(
      '<a href="#part2">jump</a><span id="part2">Part 2</span>'
    );
    const bookmark = host.querySelector('#part2')!;
    bookmark.scrollIntoView = jest.fn();
    const click = new MouseEvent('click', { bubbles: true, cancelable: true });
    host.querySelector('a')!.dispatchEvent(click);
    expect(click.defaultPrevented).toBe(true);
    expect(bookmark.scrollIntoView).toHaveBeenCalled();
  });

  it('does not reach an element with the same id outside the document', () => {
    const outside = document.createElement('div');
    outside.id = 'main';
    outside.scrollIntoView = jest.fn();
    document.body.append(outside);
    const host = render('<a href="#main">jump</a>');
    host.querySelector('a')!.click();
    expect(outside.scrollIntoView).not.toHaveBeenCalled();
  });
});
