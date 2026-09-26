import { expect, IJupyterLabPageFixture, test } from '@jupyterlab/galata';
import { Locator } from '@playwright/test';
import * as path from 'path';

const FIXTURES = path.resolve(__dirname, 'fixtures');

/**
 * Upload a fixture into the test's server folder and open it with the
 * default viewer for its file type
 */
async function open(
  page: IJupyterLabPageFixture,
  tmpPath: string,
  name: string
): Promise<Locator> {
  const target = `${tmpPath}/${name}`;
  await page.contents.uploadFile(path.join(FIXTURES, name), target);
  await page.evaluate(async (filePath: string) => {
    await (window as any).jupyterapp.commands.execute('docmanager:open', {
      path: filePath
    });
  }, target);
  return page.locator('.jp-DocReaderWidget');
}

test.describe('activation', () => {
  // Don't load JupyterLab before the test, so every log message is captured
  test.use({ autoGoto: false });

  test('should emit an activation console message', async ({ page }) => {
    const logs: string[] = [];
    page.on('console', message => {
      logs.push(message.text());
    });

    await page.goto();

    expect(
      logs.filter(
        s =>
          s ===
          'JupyterLab extension jupyterlab_doc_reader_extension is activated!'
      )
    ).toHaveLength(1);
  });
});

test.describe('DOCX', () => {
  test('renders pages in the browser with no server call and no PDF', async ({
    page,
    tmpPath
  }) => {
    const calls: string[] = [];
    page.on('request', request => {
      if (request.url().includes('jupyterlab-doc-reader-extension')) {
        calls.push(request.url());
      }
    });
    const widget = await open(page, tmpPath, 'sample.docx');

    await expect(widget.locator('section')).toHaveCount(2);
    await expect(widget).toContainText('Quarterly Report');
    await expect(widget).toContainText('Second page text.');
    await expect(widget.locator('embed, iframe')).toHaveCount(0);
    expect(calls).toEqual([]);
  });

  test('opens external links in a new tab and removes unsafe ones', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'sample.docx');

    const external = widget.locator('a', { hasText: 'Jupyter website' });
    await expect(external).toHaveAttribute('href', 'https://jupyter.org');
    await expect(external).toHaveAttribute('target', '_blank');
    const unsafe = widget.locator('a', { hasText: 'Unsafe link' });
    await expect(unsafe).toBeVisible();
    await expect(unsafe).not.toHaveAttribute('href');
    await expect(unsafe).toHaveAttribute('title', /Link disabled/);
  });

  test('does not render HTML the file embeds as an altChunk', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'altchunk.docx');

    await expect(widget).toContainText('Text around the chunk');
    await expect(widget.locator('iframe')).toHaveCount(0);
    expect(
      await page.evaluate(() => document.body.dataset.altchunk)
    ).toBeUndefined();
  });

  test('shows the renderer error for a file that is not a DOCX', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'broken.docx');

    const error = widget.locator('.jp-DocReaderWidget-error');
    await expect(error).toContainText('Cannot display this document');
    await expect(error.locator('p')).not.toBeEmpty();
  });
});

test.describe('PPTX', () => {
  const position = (page: IJupyterLabPageFixture) =>
    page.locator('.jp-DocReaderWidget-position');

  test('renders the first slide in the browser', async ({ page, tmpPath }) => {
    const widget = await open(page, tmpPath, 'sample.pptx');

    await expect(widget.locator('.jp-DocReaderWidget-stage')).toContainText(
      'Alpha slide'
    );
    await expect(position(page)).toHaveText('1 / 3');
    await expect(position(page)).toHaveAttribute('role', 'status');
    await expect(widget.locator('embed, iframe')).toHaveCount(0);
  });

  test('moves between slides with the buttons and keys', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'sample.pptx');
    const stage = widget.locator('.jp-DocReaderWidget-stage');
    await expect(position(page)).toHaveText('1 / 3');

    await page.getByTitle('Next slide').click();
    await expect(position(page)).toHaveText('2 / 3');
    await expect(stage).toContainText('Beta slide');

    await stage.click();
    const steps: [string, string][] = [
      ['ArrowRight', '3 / 3'],
      ['ArrowRight', '3 / 3'],
      ['Home', '1 / 3'],
      ['PageDown', '2 / 3'],
      ['End', '3 / 3'],
      ['PageUp', '2 / 3'],
      ['ArrowLeft', '1 / 3']
    ];
    for (const [key, expected] of steps) {
      await page.keyboard.press(key);
      await expect(position(page)).toHaveText(expected);
    }

    await page.getByTitle('Previous slide').click();
    await expect(position(page)).toHaveText('1 / 3');
  });

  test('lists every slide as a thumbnail that opens it', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'sample.pptx');
    const thumbnails = widget.locator('.jp-DocReaderWidget-thumbnail');

    await expect(thumbnails).toHaveCount(3);
    await expect(thumbnails.nth(0)).toHaveClass(/jp-mod-active/);
    await thumbnails.nth(2).click();
    await expect(position(page)).toHaveText('3 / 3');
    await expect(thumbnails.nth(2)).toHaveClass(/jp-mod-active/);
    await expect(thumbnails.nth(2)).toHaveAttribute('aria-current', 'true');
    await expect(thumbnails.nth(0)).not.toHaveClass(/jp-mod-active/);
    await expect(thumbnails.nth(0)).toHaveAttribute('aria-current', 'false');
  });

  test('zooms the slide and fits it back to the panel', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'sample.pptx');
    const slide = widget.locator('.jp-DocReaderWidget-stage > *').first();
    await expect(slide).toContainText('Alpha slide');
    const width = async () => (await slide.boundingBox())!.width;
    const fitted = await width();

    await page.getByTitle('Zoom in').click();
    await expect.poll(width).toBeGreaterThan(fitted * 1.2);
    await page.getByTitle('Zoom out').click();
    await page.getByTitle('Zoom out').click();
    await expect.poll(width).toBeLessThan(fitted * 0.8);
    await page.getByTitle('Fit slide to width').click();
    await expect.poll(width).toBeCloseTo(fitted, 0);
  });

  test('finds text on any slide and highlights it', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'sample.pptx');
    await expect(position(page)).toHaveText('1 / 3');
    const find = page.getByLabel('Find in slides');
    const status = page.locator('.jp-DocReaderWidget-findStatus');
    await expect(status).toHaveAttribute('role', 'status');

    await find.fill('zebra');
    await find.press('Enter');
    await expect(position(page)).toHaveText('3 / 3');
    await expect(status).toHaveText('1 of 1');
    await expect(widget.locator('.jp-DocReaderWidget-match')).toHaveCount(1);

    await find.fill('slide');
    await find.press('Enter');
    await expect(status).toHaveText(/^1 of \d+$/);
    await find.press('Shift+Enter');
    await expect(status).not.toHaveText(/^1 of /);

    await find.fill('no such words');
    await find.press('Enter');
    await expect(status).toHaveText('No matches');
    await expect(widget.locator('.jp-DocReaderWidget-match')).toHaveCount(0);
  });

  test('shows an error for an empty file', async ({ page, tmpPath }) => {
    const widget = await open(page, tmpPath, 'empty.pptx');

    await expect(widget.locator('.jp-DocReaderWidget-error')).toContainText(
      'The file is empty.'
    );
  });
});

test.describe('RTF', () => {
  test('renders formatted text in the browser', async ({ page, tmpPath }) => {
    const widget = await open(page, tmpPath, 'sample.rtf');

    await expect(widget).toContainText('Plain text and');
    const bold = widget.getByText('bold text');
    const weight = await bold.evaluate(
      element => getComputedStyle(element).fontWeight
    );
    expect(Number(weight)).toBeGreaterThanOrEqual(700);
  });
});

test.describe('legacy formats', () => {
  for (const [name, format, target] of [
    ['legacy.doc', 'DOC', 'DOCX'],
    ['legacy.ppt', 'PPT', 'PPTX']
  ]) {
    test(`${format} shows the unsupported message`, async ({
      page,
      tmpPath
    }) => {
      const widget = await open(page, tmpPath, name);

      const message = widget.locator('.jp-DocReaderWidget-unsupported');
      await expect(message).toContainText(`${format} files are not supported`);
      await expect(message).toContainText(`Save the file as ${target}`);
    });
  }
});

test.describe('ODT', () => {
  const frame = (widget: Locator) =>
    widget.locator('iframe.jp-DocReaderWidget-frame').contentFrame();

  test('renders formatted pages in a sandboxed frame', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'sample.odt');

    await expect(frame(widget).getByText('Sample Heading')).toBeVisible();
    await expect(frame(widget).locator('td')).toHaveCount(4);
    const bold = frame(widget).getByText('bold run');
    const weight = await bold.evaluate(
      element => getComputedStyle(element).fontWeight
    );
    expect(Number(weight)).toBeGreaterThanOrEqual(700);
    const iframe = widget.locator('iframe');
    await expect(iframe).toHaveAttribute('sandbox', 'allow-scripts');
    // an opaque origin: JupyterLab cannot reach the page, nor the page JupyterLab
    expect(
      await iframe.evaluate(element => element.contentDocument === null)
    ).toBe(true);
  });

  test('finds text and zooms with the viewer', async ({ page, tmpPath }) => {
    const widget = await open(page, tmpPath, 'sample.odt');
    const heading = frame(widget).getByText('Sample Heading');
    await expect(heading).toBeVisible();
    const find = page.getByLabel('Find in document');
    const status = page.locator('.jp-DocReaderWidget-findStatus');

    await find.fill('Cell');
    await find.press('Enter');
    await expect(status).toHaveText('1 of 4');
    await find.press('Enter');
    await expect(status).toHaveText('2 of 4');
    await find.press('Shift+Enter');
    await expect(status).toHaveText('1 of 4');
    await expect(frame(widget).locator('mark')).toHaveCount(4);
    await find.fill('no such words');
    await find.press('Enter');
    await expect(status).toHaveText('No matches');

    const height = async () => (await heading.boundingBox())!.height;
    const fitted = await height();
    await page.getByTitle('Zoom in').click();
    await expect.poll(height).toBeGreaterThan(fitted * 1.2);
    await page.getByTitle('Fit page to width').click();
    await expect.poll(height).toBeCloseTo(fitted, 0);
  });

  test('opens web links in a new tab and ignores unsafe ones', async ({
    page,
    tmpPath
  }) => {
    await page
      .context()
      .route('https://example.org/**', route => route.fulfill({ body: 'ok' }));
    const widget = await open(page, tmpPath, 'sample.odt');
    const popups: string[] = [];
    page.context().on('page', popup => popups.push(popup.url()));

    await frame(widget).getByText('Unsafe link').click();
    await frame(widget).getByText('Example link').click();
    await expect.poll(() => popups.length).toBe(1);
    await expect
      .poll(() => page.context().pages().at(-1)!.url())
      .toBe('https://example.org/');
    expect(
      await page.evaluate(() => document.body.dataset.odflink)
    ).toBeUndefined();
  });

  test('shows the viewer error for a file that is not an ODT', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'broken.odt');

    const error = widget.locator('.jp-DocReaderWidget-error');
    await expect(error).toContainText('Cannot display this document');
    await expect(error.locator('p')).not.toBeEmpty();
  });
});

test.describe('ODP', () => {
  test('renders every slide and finds text across them', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'sample.odp');
    const frame = widget.locator('iframe').contentFrame();

    for (const title of ['Alpha', 'Beta', 'Gamma']) {
      await expect(frame.getByText(title, { exact: true })).toBeAttached();
    }
    const find = page.getByLabel('Find in slides');
    await find.fill('zebra');
    await find.press('Enter');
    await expect(page.locator('.jp-DocReaderWidget-findStatus')).toHaveText(
      '1 of 1'
    );
    await expect(frame.locator('mark')).toHaveCount(1);
  });
});

test.describe('XLSX', () => {
  test('draws the workbook with its sheet tabs and finds cells', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'sample.xlsx');
    const status = page.locator('.jp-DocReaderWidget-findStatus');

    // the viewer draws before its load resolves; find works after that
    await expect(widget.locator('.jp-DocReaderWidget-loading')).toHaveCount(0, {
      timeout: 30000
    });
    await expect(widget.locator('canvas').first()).toBeVisible();
    await expect(widget.getByText('Data', { exact: true })).toBeVisible();
    await expect(widget.getByText('Other', { exact: true })).toBeVisible();
    await expect(widget.locator('embed, iframe')).toHaveCount(0);

    const find = page.getByLabel('Find in sheets');
    await find.fill('Pears');
    await find.press('Enter');
    await expect(status).toHaveText('1 of 1');
    await find.fill('Second sheet cell');
    await find.press('Enter');
    await expect(status).toHaveText('1 of 1');
    await find.fill('no such words');
    await find.press('Enter');
    await expect(status).toHaveText('No matches');
  });

  test('shows the viewer error for a file that is not an XLSX', async ({
    page,
    tmpPath
  }) => {
    const widget = await open(page, tmpPath, 'broken.xlsx');

    await expect(widget.locator('.jp-DocReaderWidget-error')).toContainText(
      'Cannot display this document'
    );
  });
});
